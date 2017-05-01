VERSION 5.00
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form AGC2432C 
   Caption         =   "理化委托单-PWHT_AGC2432C"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   WindowState     =   2  'Maximized
   Begin VB.TextBox TXT_NO 
      BackColor       =   &H00FFFFFF&
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
      Left            =   6840
      TabIndex        =   21
      Tag             =   "plt"
      Top             =   270
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.TextBox TXT_SIZE 
      BackColor       =   &H00FFFFFF&
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
      Left            =   5490
      TabIndex        =   20
      Tag             =   "plt"
      Top             =   270
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.TextBox PLT_NAME 
      BackColor       =   &H00FFFFFF&
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
      Left            =   8220
      TabIndex        =   19
      Tag             =   "plt"
      Top             =   270
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CheckBox txt_check 
      BackColor       =   &H00E0E0E0&
      Caption         =   "已处理对象"
      Height          =   255
      Left            =   5190
      TabIndex        =   13
      Top             =   30
      Width           =   1215
   End
   Begin VB.TextBox txt_plt 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1245
      TabIndex        =   12
      Tag             =   "plt"
      Top             =   465
      Width           =   540
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   435
      Left            =   1815
      TabIndex        =   9
      Top             =   360
      Width           =   2400
      Begin VB.CheckBox txt_DH_FL 
         BackColor       =   &H00E0E0E0&
         Caption         =   "热处理"
         Height          =   240
         Left            =   45
         TabIndex        =   11
         Top             =   150
         Width           =   915
      End
      Begin VB.TextBox txt_line 
         Height          =   300
         Left            =   1785
         MaxLength       =   10
         TabIndex        =   10
         Top             =   100
         Width           =   450
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Left            =   1110
         Top             =   100
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   556
         Caption         =   "产线"
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
   End
   Begin VB.TextBox txt_smp_sent_no 
      Height          =   315
      Left            =   9690
      MaxLength       =   13
      TabIndex        =   8
      Top             =   465
      Width           =   1440
   End
   Begin VB.TextBox TXT_CUT_PLT 
      Height          =   315
      Left            =   11190
      TabIndex        =   7
      Top             =   465
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.CommandButton CmdSEND 
      Caption         =   "发送委托信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   11415
      TabIndex        =   6
      Top             =   360
      Width           =   1725
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "打印委托单"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   13380
      TabIndex        =   5
      Top             =   360
      Width           =   1725
   End
   Begin VB.TextBox txt_smp_no 
      Height          =   315
      Left            =   5820
      MaxLength       =   14
      TabIndex        =   4
      Top             =   465
      Width           =   1650
   End
   Begin VB.CheckBox txt_smp_fl 
      BackColor       =   &H00E0E0E0&
      Caption         =   "复样"
      Height          =   255
      Left            =   4470
      TabIndex        =   3
      Top             =   30
      Width           =   735
   End
   Begin VB.ComboBox txt_HTM_METH 
      Height          =   300
      ItemData        =   "AGC2432C.frx":0000
      Left            =   13170
      List            =   "AGC2432C.frx":000D
      TabIndex        =   2
      Top             =   0
      Width           =   1065
   End
   Begin VB.CheckBox Chc_OutOrder 
      BackColor       =   &H00E0E0E0&
      Caption         =   "可委托出口船板"
      Height          =   255
      Left            =   6510
      TabIndex        =   1
      Top             =   30
      Width           =   1695
   End
   Begin VB.TextBox Txt_OutOrder 
      Height          =   270
      Left            =   11310
      TabIndex        =   0
      Text            =   "0"
      Top             =   30
      Visible         =   0   'False
      Width           =   375
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8295
      Left            =   -120
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   840
      Width           =   15180
      _Version        =   393216
      _ExtentX        =   26776
      _ExtentY        =   14631
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
      MaxCols         =   94
      MaxRows         =   1
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      SpreadDesigner  =   "AGC2432C.frx":001A
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   30
      Top             =   0
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "日期"
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
   Begin InDate.ULabel ULabel7 
      Height          =   315
      Left            =   15
      Top             =   465
      Width           =   1215
      _ExtentX        =   2143
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
      ForeColor       =   0
   End
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   8310
      Top             =   465
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "委托单号"
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
   Begin InDate.ULabel ULabel8 
      Height          =   315
      Left            =   8310
      Top             =   0
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   556
      Caption         =   "要求完成时间"
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
   Begin InDate.UDate dtp_end_date 
      Height          =   315
      Left            =   9735
      TabIndex        =   15
      Top             =   0
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
   Begin InDate.ULabel ULabel9 
      Height          =   315
      Left            =   4470
      Top             =   465
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "试样号"
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
      ForeColor       =   0
   End
   Begin InDate.UDate dtp_prod_fr 
      Height          =   315
      Left            =   1260
      TabIndex        =   16
      Top             =   0
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
      MaxLength       =   10
   End
   Begin InDate.UDate dtp_prod_to 
      Height          =   315
      Left            =   2730
      TabIndex        =   17
      Top             =   0
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
      MaxLength       =   10
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   11790
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "热处理方式"
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
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "每炉成分委托数量系统自动管控，请勿再人工勾选"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   8160
      TabIndex        =   18
      Top             =   480
      Width           =   4965
   End
End
Attribute VB_Name = "AGC2432C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       Template System
'-- Sub_System Name   Common
'-- Program Name      Refer Template
'-- Program ID        Refer
'-- Document No       Q-00-0010(Specification)
'-- Designer          zhang lin
'-- Coder             zhang lin
'-- Date              2007.3.22
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

Dim pColumn1 As New Collection      'Spread Primary Key Collection
Dim nColumn1 As New Collection      'Spread necessary Column Collection
Dim mColumn1 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn1 As New Collection      'Spread Insert Column Collection
Dim aColumn1 As New Collection      'Master -> Spread Column Collection
Dim lColumn1 As New Collection      'Spread Lock Column Collection


Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

'Dim iSumCol As New Collection       'Sum Column

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2
'---------------------------------------------------------------------------------------------
'------------------------------ Report Variable ----------------------------------------------
'---------------------------------------------------------------------------------------------
Dim xlApp       As Object
Dim xlSheet     As Object

Dim arrRecords1 As Variant

Dim sQuery      As String
Dim sErrMsg     As String
Dim sDate       As String
Dim AdoRs       As ADODB.Recordset

Const SS1_REQ_END_DATE = 48
Const SS1_SMP_CUT_PLT = 51
Const SS1_SMP_NO = 49
Const SS1_SMP_NO_ZW = 50
Const SS1_SMPNO = 2
Const SS1_STDSPEC = 4
Const SS1_SIZE = 5
Const SS1_NO = 6
Const SS1_LSA = 7
Const SS1_LSB = 8
Const SS1_LSC = 9
Const SS1_LSD = 10
Const SS1_LSE = 11
Const SS1_LSF = 12
Const SS1_LSG = 13
Const SS1_LSH = 14
Const SS1_LSI = 15
Const SS1_LSJ = 16
Const SS1_LSK = 17
Const SS1_LSL = 18
Const SS1_WQA = 19
Const SS1_CJA = 20
Const SS1_CJB = 21
Const SS1_CJC = 22
Const SS1_CJD = 23
Const SS1_CJE = 24
Const SS1_CJF = 25
Const SS1_CJG = 26
Const SS1_CJH = 27
Const SS1_CJI = 28
Const SS1_CJJ = 29
Const SS1_CJK = 30
Const SS1_CJL = 31
Const SS1_YDA = 32
Const SS1_ZXA = 33
Const SS1_ZXB = 34
Const SS1_ZXC = 35
Const SS1_ZXD = 36
Const SS1_ZXE = 37
Const SS1_ZXF = 38
Const SS1_ZXG = 39
Const SS1_ZXH = 40
Const SS1_ZXI = 41
Const SS1_ZXJ = 42
Const SS1_ZXK = 43
Const SS1_ZXL = 44
Const SS1_JXA = 45
Const SS1_JZA = 46
Const SS1_TEST = 47
Const SS1_LOC = 53
Const SS1_SMP_CD = 54
Const SS1_CHECK = 1
Const SS1_LA_SMP_CD = 55
Const SS1_LB_SMP_CD = 56
Const SS1_LC_SMP_CD = 57
Const SS1_LD_SMP_CD = 58
Const SS1_LE_SMP_CD = 59
Const SS1_LF_SMP_CD = 60
Const SS1_LG_SMP_CD = 61
Const SS1_LH_SMP_CD = 62
Const SS1_LI_SMP_CD = 63
Const SS1_LJ_SMP_CD = 64
Const SS1_LK_SMP_CD = 65
Const SS1_LL_SMP_CD = 66
Const SS1_WQ_SMP_CD = 67
Const SS1_CA_SMP_CD = 68
Const SS1_CB_SMP_CD = 69
Const SS1_CC_SMP_CD = 70
Const SS1_CD_SMP_CD = 71
Const SS1_CE_SMP_CD = 72
Const SS1_CF_SMP_CD = 73
Const SS1_CG_SMP_CD = 74
Const SS1_CH_SMP_CD = 75
Const SS1_CI_SMP_CD = 76
Const SS1_CJ_SMP_CD = 77
Const SS1_CK_SMP_CD = 78
Const SS1_CL_SMP_CD = 79
Const SS1_YD_SMP_CD = 80
Const SS1_ZA_SMP_CD = 81
Const SS1_ZB_SMP_CD = 82
Const SS1_ZC_SMP_CD = 83
Const SS1_ZD_SMP_CD = 84
Const SS1_ZE_SMP_CD = 85
Const SS1_ZF_SMP_CD = 86
Const SS1_ZG_SMP_CD = 87
Const SS1_ZH_SMP_CD = 88
Const SS1_ZI_SMP_CD = 89
Const SS1_ZJ_SMP_CD = 90
Const SS1_ZK_SMP_CD = 91
Const SS1_ZL_SMP_CD = 92
Const SS1_XA_SMP_CD = 93
Const SS1_JA_SMP_CD = 94


Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"
         
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
     Call Gp_Ms_Collection(DTP_PROD_FR, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(DTP_PROD_TO, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'    Call Gp_Ms_Collection(TXT_SLAB_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_PLT, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_line, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(txt_smp_sent_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_check, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_DH_FL, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_SMP_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_SMP_FL, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_htm_meth, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(Txt_OutOrder, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '添加出口查询20130328wch
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
    Call Gp_Sp_Collection(ss1, 2, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 35, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 36, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 37, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 38, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 39, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 40, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 41, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 42, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 43, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 44, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 45, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 46, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 47, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 48, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 49, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 50, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 51, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 52, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 53, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 54, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 55, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 56, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 57, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 58, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 59, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 60, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 61, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 62, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 63, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 64, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 65, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 66, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 67, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 68, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 69, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 70, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 71, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 72, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 73, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 74, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 75, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 76, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 77, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 78, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 79, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 80, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 81, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 82, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 83, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 84, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 85, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 86, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 87, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 88, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 89, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 90, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 91, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 92, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 93, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 94, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
'    sc1.Add Item:="AGC2432C.P_ONEROW", Key:="P-O"
    sc1.Add Item:="AGC2432C.P_REFER", Key:="P-R"
    sc1.Add Item:="AGC2432C.P_MODIFY", Key:="P-M"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxRows, Key:="Last"
    
    Proc_Sc.Add Item:=sc1, Key:="Sc"
'    Call Gp_Sp_ColHidden(ss1, 27, True)
'    Call Gp_Sp_ColHidden(ss1, 34, True)
'    Call Gp_Sp_ColHidden(ss1, 31, True)
'    Call Gp_Sp_ColHidden(ss1, 32, True)
'    Call Gp_Sp_ColHidden(ss1, 33, True)
'    Call Gp_Sp_ColHidden(ss1, 9, True)
    
    Me.KeyPreview = True

End Sub


Private Sub Chc_OutOrder_Click()
 If Chc_OutOrder.Value = ssCBChecked Then
        Txt_OutOrder = "1"
    Else
        Txt_OutOrder = "0"
    End If
End Sub

Private Sub cmdReport_Click()

    Dim sQuery As String
    Dim arrRecords1 As Variant
    Dim AdoRs As ADODB.Recordset
    
'    If ss1.MaxRows < 1 Then Exit Sub
    
    If txt_smp_sent_no.Text = "" Then
       Call MsgBox("请输入委托单号，再打印！", vbCritical, "系统提示信息")
       Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
       
    Set AdoRs = New ADODB.Recordset
    
       
    Call ExcelPrn_Pile
    
    Call PRINT_Click
    
End Sub

Private Sub PRINT_Click()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(2, 4) As Variant
    Dim ret_Result_ErrCode As Integer
    Dim ret_Result_ErrMsg As String
    Dim iRow, iCnt As Integer
    Dim sQuery As String
    Dim adoCmd As ADODB.Command

    OutParam(1, 1) = "arg_e_code"
    OutParam(1, 2) = adInteger
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 1

    'Return Error Messsage Parameter
    OutParam(2, 1) = "arg_e_msg"
    OutParam(2, 2) = adVarChar
    OutParam(2, 3) = adParamOutput
    OutParam(2, 4) = 256
    
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command

    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1

    sQuery = "{call AGC2432C.P_MODIFY2 ( '" + txt_smp_sent_no.Text + "',?,? )}"
    
    adoCmd.CommandText = sQuery

    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(2, 1), OutParam(2, 2), OutParam(2, 3), OutParam(2, 4))
    
    adoCmd.Execute , , adExecuteNoRecords

    'Process Error Check
    If adoCmd("arg_e_msg") <> "0" Then
        ret_Result_ErrCode = adoCmd("arg_e_code")
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        
        sErrMessg = "Error Code : " & ret_Result_ErrCode & vbCrLf & "Error Mesg : " & ret_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
        Set adoCmd = Nothing
        Exit Sub
    End If

    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    
    Err.Raise Err.Number, Err.Description & sQuery
    
    
End Sub
        


Private Sub CmdSEND_Click()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(2, 4) As Variant
    Dim ret_Result_ErrCode As String
    Dim ret_Result_ErrMsg As String
    Dim iRow, iCnt As Integer
    Dim sQuery As String
    Dim adoCmd As ADODB.Command
    
    If ss1.MaxRows < 1 Then Exit Sub
  
    If txt_smp_sent_no.Text = "" Then
       Call MsgBox("请输入委托单号，再打印！", vbCritical, "系统提示信息")
       Exit Sub
    End If
    

    Screen.MousePointer = vbHourglass
    
    'Return Error Code Parameter
    OutParam(1, 1) = "arg_e_code"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 2

    'Return Error Messsage Parameter
    OutParam(2, 1) = "arg_e_msg"
    OutParam(2, 2) = adVarChar
    OutParam(2, 3) = adParamOutput
    OutParam(2, 4) = 256
    
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command

    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1

    sQuery = "{call AQC1061P ('" + txt_smp_sent_no.Text + "','',?,?)}"
    
'    Debug.Print sQuery
    

    adoCmd.CommandText = sQuery

    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(2, 1), OutParam(2, 2), OutParam(2, 3), OutParam(2, 4))
    
    adoCmd.Execute , , adExecuteNoRecords

    'Process Error Check
    If adoCmd("arg_e_msg") <> "YY" Then
        ret_Result_ErrCode = adoCmd("arg_e_code")
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        
        sErrMessg = "Error Code : " & ret_Result_ErrCode & vbCrLf & "Error Mesg : " & ret_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
        Set adoCmd = Nothing
        Exit Sub
    Else
       Call MsgBox("委托单'" + txt_smp_sent_no.Text + "'已发送！", vbOKOnly, "系统提示信息")
    End If

    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    
    Err.Raise Err.Number, Err.Description & sQuery
    
End Sub

Private Sub dtp_prod_fr_DblClick()
    DTP_PROD_FR.RawData = Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYYMMDD') FROM DUAL")
End Sub

Private Sub dtp_prod_to_DblClick()
    DTP_PROD_TO.RawData = Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYYMMDD') FROM DUAL")
End Sub

Private Sub Form_Activate()
    
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    Call subButtonHide
        
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
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    
    Call Gp_Ms_Cls(Mc1("rControl"))

    Call Gp_Ms_NeceColor(Mc1("nControl"))
   
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_HdColColor(ss1, 1)
    
    DTP_PROD_FR.RawData = Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYYMMDD') FROM DUAL")
    DTP_PROD_TO.RawData = Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYYMMDD') FROM DUAL")

    If App.Title = "BG" Then
       TXT_PLT = "C1"
    ElseIf App.Title = "CG" Then
       TXT_PLT = "C3"
    End If
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "Q-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault
    
    txt_SMP_FL.Value = 0
    txt_check.Value = 0
    Call subButtonHide
    txt_smp_sent_no.Text = ""
    TXT_CUT_PLT.Text = ""
    PLT_NAME.Text = "板卷厂"
    
    Call Gp_Sp_ColHidden(ss1, 55, True)
    Call Gp_Sp_ColHidden(ss1, 56, True)
    Call Gp_Sp_ColHidden(ss1, 57, True)
    Call Gp_Sp_ColHidden(ss1, 58, True)
    Call Gp_Sp_ColHidden(ss1, 59, True)
    Call Gp_Sp_ColHidden(ss1, 60, True)
    Call Gp_Sp_ColHidden(ss1, 61, True)
    Call Gp_Sp_ColHidden(ss1, 62, True)
    Call Gp_Sp_ColHidden(ss1, 63, True)
    Call Gp_Sp_ColHidden(ss1, 64, True)
    Call Gp_Sp_ColHidden(ss1, 65, True)
    Call Gp_Sp_ColHidden(ss1, 66, True)
    Call Gp_Sp_ColHidden(ss1, 67, True)
    Call Gp_Sp_ColHidden(ss1, 68, True)
    Call Gp_Sp_ColHidden(ss1, 69, True)
    Call Gp_Sp_ColHidden(ss1, 70, True)
    Call Gp_Sp_ColHidden(ss1, 71, True)
    Call Gp_Sp_ColHidden(ss1, 72, True)
    Call Gp_Sp_ColHidden(ss1, 73, True)
    Call Gp_Sp_ColHidden(ss1, 74, True)
    Call Gp_Sp_ColHidden(ss1, 75, True)
    Call Gp_Sp_ColHidden(ss1, 76, True)
    Call Gp_Sp_ColHidden(ss1, 77, True)
    Call Gp_Sp_ColHidden(ss1, 78, True)
    Call Gp_Sp_ColHidden(ss1, 79, True)
    Call Gp_Sp_ColHidden(ss1, 80, True)
    Call Gp_Sp_ColHidden(ss1, 81, True)
    Call Gp_Sp_ColHidden(ss1, 82, True)
    Call Gp_Sp_ColHidden(ss1, 83, True)
    Call Gp_Sp_ColHidden(ss1, 84, True)
    Call Gp_Sp_ColHidden(ss1, 85, True)
    Call Gp_Sp_ColHidden(ss1, 86, True)
    Call Gp_Sp_ColHidden(ss1, 87, True)
    Call Gp_Sp_ColHidden(ss1, 88, True)
    Call Gp_Sp_ColHidden(ss1, 89, True)
    Call Gp_Sp_ColHidden(ss1, 90, True)
    Call Gp_Sp_ColHidden(ss1, 91, True)
    Call Gp_Sp_ColHidden(ss1, 92, True)
    Call Gp_Sp_ColHidden(ss1, 93, True)
    Call Gp_Sp_ColHidden(ss1, 94, True)
    
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
    Set sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")


End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    End If
    
    txt_SMP_FL.Value = 0
    txt_check.Value = 0
    Call subButtonHide
    txt_smp_sent_no.Text = ""
    TXT_CUT_PLT.Text = ""
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub
Public Sub Form_Ref()

On Error GoTo Refer_Err

    Dim SMESG As String
    Dim iRow As Integer
    Dim iCol As Integer
    Dim sUrgnt_Fl As String
    Dim OutOrder As String
    Dim iCount As Integer
    Dim sQuery As String
    Dim arrRecords1 As Variant
    Dim AdoRs As ADODB.Recordset
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
                
        If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
            ss1.OperationMode = OperationModeNormal
            
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call subButtonHide
            
            If ss1.MaxRows >= 1 Then
               ss1.Row = 1
               ss1.Col = SS1_SMP_CUT_PLT
               TXT_CUT_PLT.Text = ss1.Text
            End If
            
            Exit Sub
        End If
            
    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Pro()

    Dim AdoRs       As ADODB.Recordset
    Dim sQuery      As String
    Dim i           As Integer
    Dim sREQ_DATE   As String
    Dim COUNT1   As Integer
    Dim sStdspec As String

    COUNT1 = 0

    If txt_check.Value = "0" Then
    
        Set AdoRs = New ADODB.Recordset
           
        sQuery = "SELECT Gf_SMP_SEND_NO( "
        sQuery = sQuery & "'" & TXT_PLT & "') "
        sQuery = sQuery & "FROM DUAL"
                      
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
        If Not AdoRs.BOF And Not AdoRs.EOF Then
           For i = 1 To ss1.MaxRows
               ss1.Row = i
               
               ss1.Col = 0
               If ss1.Text = "Update" Or ss1.Text = "Insert" Then
               
                  ss1.Col = 3
                  ss1.Text = AdoRs.Fields(0) & ""
                  
                  ss1.Col = SS1_REQ_END_DATE
                  sREQ_DATE = ss1.Text
                  If sREQ_DATE = "" Then
                     ss1.Text = dtp_end_date.RawData
                  End If
                  
               End If
               
               ss1.Col = 1
              If ss1.Text = "1" Then
                 COUNT1 = COUNT1 + 1
                 If COUNT1 > 2 Then
                    Call MsgBox("一张委托单不能超过2个试样！", vbCritical, "系统提示信息")
                    Exit Sub
                 End If
              End If
             
              ss1.Col = SS1_STDSPEC
              sStdspec = ss1.Text
              If sStdspec <> ss1.Text Then
                 Call Gp_MsgBoxDisplay("不一样钢种")
                 Exit Sub
              End If
              
           Next i
        End If
        AdoRs.Close
        Set AdoRs = Nothing
        
    End If
 
    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
      Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
      Call subButtonHide
    End If
    
End Sub

Public Sub Form_Ins()
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 27)
    ss1.SetFocus
    

End Sub

Public Sub Spread_Can()

    
    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
    

End Sub

Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 27)
    
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

Public Sub Form_Exit()

    Unload Me
    
End Sub
Public Sub Spread_Del()
    
    Call Gp_Sp_Del(Proc_Sc("SC"))

End Sub



Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    

End Sub


Private Sub ss1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

    Dim sCheck As Integer
    
    Dim iCol As Long
    Dim iRow As Long
    Dim iMode As Integer
        
    iCol = Col
    iRow = Row

    If Row <= 0 Then Exit Sub
    If Not Gf_Sc_Authority(sAuthority, "U") Then Exit Sub
    
    ss1.Row = iRow
    
    ss1.Col = 0
    ss1.Text = "Update"
    
    ss1.Col = SS1_CHECK
    sCheck = ss1.Value
    
    If sCheck = 0 Then
       ss1.Col = 0
       ss1.Text = ""
    End If
    
    ss1.Col = SS1_SMP_NO
    ss1.Text = sUserID
        
End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)

    If txt_check.Value = 1 Then

        ss1.Row = ss1.ActiveRow
        ss1.Col = 3
        txt_smp_sent_no.Text = ss1.Text
        ss1.Col = SS1_SMP_CUT_PLT
        TXT_CUT_PLT.Text = ss1.Text
        ss1.Col = SS1_SIZE
        TXT_SIZE.Text = ss1.Text
        ss1.Col = SS1_NO
        txt_no.Text = ss1.Text
    End If
    
End Sub


Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 27)
    
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

'Private Sub ExcelPrn()
'    Dim I               As Integer
'    Dim xlApp           As Object
'    Dim xlSheet         As Object
'    Dim sRow            As String
'
'    If ss1.MaxRows < 1 Then
'       MsgBox "请先查询数据再打印！", vbCritical, "系统提示信息"
'       Exit Sub
'    End If
'
''    If Trim(TXT_CHECK) <> dtp_yy_mm.RawData Then
''       MsgBox "选择的日期没有进行查询，请先查询数据再打印！", vbCritical, "系统提示信息"
''       Exit Sub
''
''    End If
'
'    Screen.MousePointer = vbHourglass
'
'    On Error Resume Next
'
'    Set xlApp = GetObject(, "Excel.Application")
'    If Err.Number <> 0 Then
'        Set xlApp = CreateObject("Excel.Application")
'    End If
'
'    Err.Clear
'
'    xlApp.Workbooks.Open (App.Path & "\AGC2401C.xls")
'
'    Set xlSheet = xlApp.Worksheets("Sheet1")
'    xlApp.Sheets("Sheet1").Select
'
'    xlApp.Range("B3").Value = txt_smp_sent_no.Text
'    xlApp.Range("B4").Value = TXT_CUT_PLT.Text
'    xlApp.Range("C4").Value = Mid(Now, 1, 4) + "年" + _
'                              Mid(Now, 6, 2) + "月" + _
'                              Mid(Now, 9, 2) + "日"
'
'    Clipboard.Clear
'    ss1.Row = 1: ss1.Col = 3: ss1.Row2 = ss1.MaxRows: ss1.Col2 = 21
'    Clipboard.SetText ss1.Clip
'    xlApp.Range("A7").Select
'    xlApp.ActiveSheet.Paste
'
'
'    Clipboard.Clear
'
''    sRow = "A" & ss1.MaxRows + 9 & ":B" & ss1.MaxRows + 6
''    xlApp.Range(sRow).MERGECELLS = True
'    sRow = "A" & ss1.MaxRows + 7
'    xlApp.Range(sRow).Value = "委托人：" & sUsername
'    xlApp.Range(sRow).Font.Size = 10
'    sRow = "B" & ss1.MaxRows + 7
'    xlApp.Range(sRow).Value = "委托时间：" & Format(Now, "YYYY-MM-DD HH:MM:SS")
'    xlApp.Range(sRow).Font.Size = 10
''    xlApp.Range("O5").Value = Format(Now, "YYYY-MM-DD HH:MM:SS")
'    sRow = "F" & ss1.MaxRows + 7
'    xlApp.Range(sRow).Value = "送样人："
'    xlApp.Range(sRow).Font.Size = 10
''    sRow = "K" & ss1.MaxRows + 7
''    xlApp.Range(sRow).Value = "收样人："
''    xlApp.Range(sRow).Font.Size = 10
'    sRow = "M" & ss1.MaxRows + 7
'    xlApp.Range(sRow).Value = "送样时间："
'    xlApp.Range(sRow).Font.Size = 10
''    xlApp.Range(sRow).Font.Bold = True
'    xlApp.ActiveSheet.Paste
'
'
'    ss1.ClearSelection
'    With xlApp.Application.FindFormat.Borders
'        .LineStyle = 1
'    End With
'
'    sRow = "A7:S" & ss1.MaxRows + 6
'    xlApp.Range(sRow).Select
'    With xlApp.Selection.Borders
'        .LineStyle = 1
'    End With
''    xlApp.Columns("C:E").AutoFit
''    xlApp.Columns("J").AutoFit
'    Screen.MousePointer = vbDefault
'    xlApp.Application.Visible = True
''    xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
''    xlApp.DisplayAlerts = False
''    xlSheet.Close
'
''    Set xlSheet = Nothing
''    Set xlApp = Nothing
'
'    Exit Sub
'
'ErrHandle:
'    MsgBox Error
'    Set xlSheet = Nothing
'    Set xlApp = Nothing
'    Screen.MousePointer = vbDefault
'End Sub

Private Sub subButtonHide()

Dim iRow, iCol As Integer


    If txt_check.Value = 1 Then
        MDIMain.MenuTool.Buttons(8).Enabled = True    'Row delete
        MDIMain.MenuTool.Buttons(7).Enabled = False    'Row insert
        cmdReport.Visible = True
        CmdSEND.Visible = True
        ULabel6.Visible = True
        txt_smp_sent_no.Visible = True
        ULabel8.Visible = False
        dtp_end_date.Visible = False
        Label1.Visible = False
        
        Call Gp_Sp_ColHidden(ss1, 3, False)
        Call Gp_Sp_BlockLock(ss1, 1, ss1.MaxCols, 1, ss1.MaxRows, True)
        
    Else
        MDIMain.MenuTool.Buttons(8).Enabled = False    'Row delete
        MDIMain.MenuTool.Buttons(7).Enabled = False    'Row insert
        cmdReport.Visible = False
        CmdSEND.Visible = False
        ULabel6.Visible = False
        txt_smp_sent_no.Visible = False
        ULabel8.Visible = True
        dtp_end_date.Visible = True
        Label1.Visible = True

        Call Gp_Sp_ColHidden(ss1, 3, True)
    End If

End Sub

Private Sub txt_check_Click()

    Call subButtonHide
    Call Form_Ref
    txt_smp_sent_no.Text = ""
    
End Sub

Private Sub txt_smp_fl_Click()

    Call Form_Ref
    
End Sub

Private Sub TXT_PLT_Change()
    If TXT_PLT.Text = "C1" Then
       PLT_NAME.Text = "板卷厂"
    ElseIf TXT_PLT.Text = "C2" Then
       PLT_NAME.Text = "宽厚板厂"
    ElseIf TXT_PLT.Text = "C3" Then
       PLT_NAME.Text = "中板厂"
    End If
End Sub
Private Sub txt_plt_KeyUp(KeyCode As Integer, Shift As Integer)

'    If KeyCode = vbKeyF4 Then
'
'        DD.sWitch = "MS"
'        DD.sKey = "C0001"
'        DD.rControl.Add Item:=txt_plt
'
'        DD.nameType = "2"
'
'        Call Gf_Common_DD(M_CN1, KeyCode)
'
'        Exit Sub
'
'    End If

End Sub


Private Sub ExcelPrn_Pile()

    Dim i               As Integer
    Dim xlApp           As Object
    Dim xlSheet         As Object
    Dim sRow            As String
    Dim sSmpNo          As String
    
    Dim sText           As String
    Dim sSmpcd          As String
    
    Dim sSize           As String
    Dim sNo             As String
    
    If ss1.MaxRows < 1 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
     
    On Error Resume Next
    
    Set xlApp = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then
        Set xlApp = CreateObject("Excel.Application")
    End If
    
    Err.Clear

    xlApp.Workbooks.Open (App.Path & "\AGC2432C.xlsx")
    
    Set xlSheet = xlApp.Worksheets("Sheet1")
    xlApp.Sheets("Sheet1").Select
    
    sSmpNo = txt_smp_sent_no.Text
    
    xlApp.Range("A2").Value = "委托单号：" & txt_smp_sent_no.Text & "           工厂：" & PLT_NAME.Text
    
    ss1.Row = 1
    ss1.Col = SS1_SMPNO:         xlApp.Range("A4").Value = ss1.Text
    ss1.Col = SS1_STDSPEC:       xlApp.Range("B4").Value = ss1.Text
'    ss1.Col = SS1_SIZE:          xlApp.Range("C4").Value = ss1.Text
'    xlApp.Range("C4").Value = ss1.Col = SS1_SIZE & ",数量：" & ss1.Col = SS1_NO
'    xlApp.Range("C4").Value = TXT_SIZE.Text & ",数量：" & TXT_NO.Text
    With ss1
        .Row = 1:   .Col = SS1_SIZE
        sSize = .Text
        .Row = 1:   .Col = SS1_NO
        sNo = .Text
        xlApp.Range("C4").Value = sSize & ",数量：" & sNo
    End With
    ss1.Col = SS1_LSA:           xlApp.Range("D4").Value = ss1.Text
    ss1.Col = SS1_LA_SMP_CD:           xlApp.Range("E4").Value = ss1.Text
    ss1.Col = SS1_LSB:           xlApp.Range("F4").Value = ss1.Text
    ss1.Col = SS1_LB_SMP_CD:           xlApp.Range("G4").Value = ss1.Text
    ss1.Col = SS1_LSC:           xlApp.Range("H4").Value = ss1.Text
    ss1.Col = SS1_LC_SMP_CD:           xlApp.Range("I4").Value = ss1.Text
    ss1.Col = SS1_LSD:           xlApp.Range("J4").Value = ss1.Text
    ss1.Col = SS1_LD_SMP_CD:           xlApp.Range("K4").Value = ss1.Text
    ss1.Col = SS1_LSE:           xlApp.Range("L4").Value = ss1.Text
    ss1.Col = SS1_LE_SMP_CD:           xlApp.Range("M4").Value = ss1.Text
    ss1.Col = SS1_LSF:           xlApp.Range("N4").Value = ss1.Text
    ss1.Col = SS1_LF_SMP_CD:           xlApp.Range("O4").Value = ss1.Text
    ss1.Col = SS1_LSG:           xlApp.Range("P4").Value = ss1.Text
    ss1.Col = SS1_LG_SMP_CD:           xlApp.Range("Q4").Value = ss1.Text
    ss1.Col = SS1_LSH:           xlApp.Range("R4").Value = ss1.Text
    ss1.Col = SS1_LH_SMP_CD:           xlApp.Range("S4").Value = ss1.Text
    ss1.Col = SS1_LSI:           xlApp.Range("T4").Value = ss1.Text
    ss1.Col = SS1_LI_SMP_CD:           xlApp.Range("U4").Value = ss1.Text
    ss1.Col = SS1_LSJ:           xlApp.Range("V4").Value = ss1.Text
    ss1.Col = SS1_LJ_SMP_CD:           xlApp.Range("W4").Value = ss1.Text
    ss1.Col = SS1_LSK:           xlApp.Range("X4").Value = ss1.Text
    ss1.Col = SS1_LK_SMP_CD:           xlApp.Range("Y4").Value = ss1.Text
    ss1.Col = SS1_LSL:           xlApp.Range("Z4").Value = ss1.Text
    ss1.Col = SS1_LL_SMP_CD:           xlApp.Range("AA4").Value = ss1.Text
    
    ss1.Col = SS1_CJA:           xlApp.Range("D6").Value = ss1.Text
    ss1.Col = SS1_CA_SMP_CD:           xlApp.Range("E6").Value = ss1.Text
    ss1.Col = SS1_CJB:           xlApp.Range("F6").Value = ss1.Text
    ss1.Col = SS1_CB_SMP_CD:           xlApp.Range("G6").Value = ss1.Text
    ss1.Col = SS1_CJC:           xlApp.Range("H6").Value = ss1.Text
    ss1.Col = SS1_CC_SMP_CD:           xlApp.Range("I6").Value = ss1.Text
    ss1.Col = SS1_CJD:           xlApp.Range("J6").Value = ss1.Text
    ss1.Col = SS1_CD_SMP_CD:           xlApp.Range("K6").Value = ss1.Text
    ss1.Col = SS1_CJE:           xlApp.Range("L6").Value = ss1.Text
    ss1.Col = SS1_CE_SMP_CD:           xlApp.Range("M6").Value = ss1.Text
    ss1.Col = SS1_CJF:           xlApp.Range("N6").Value = ss1.Text
    ss1.Col = SS1_CF_SMP_CD:           xlApp.Range("O6").Value = ss1.Text
    ss1.Col = SS1_CJG:           xlApp.Range("P6").Value = ss1.Text
    ss1.Col = SS1_CG_SMP_CD:           xlApp.Range("Q6").Value = ss1.Text
    ss1.Col = SS1_CJH:           xlApp.Range("R6").Value = ss1.Text
    ss1.Col = SS1_CH_SMP_CD:           xlApp.Range("S6").Value = ss1.Text
    ss1.Col = SS1_CJI:           xlApp.Range("T6").Value = ss1.Text
    ss1.Col = SS1_CI_SMP_CD:           xlApp.Range("U6").Value = ss1.Text
    ss1.Col = SS1_CJJ:           xlApp.Range("V6").Value = ss1.Text
    ss1.Col = SS1_CJ_SMP_CD:           xlApp.Range("W6").Value = ss1.Text
    ss1.Col = SS1_CJK:           xlApp.Range("X6").Value = ss1.Text
    ss1.Col = SS1_CK_SMP_CD:           xlApp.Range("Y6").Value = ss1.Text
    ss1.Col = SS1_CJL:           xlApp.Range("Z6").Value = ss1.Text
    ss1.Col = SS1_CL_SMP_CD:           xlApp.Range("AA6").Value = ss1.Text
    
    ss1.Col = SS1_ZXA:           xlApp.Range("D8").Value = ss1.Text
    ss1.Col = SS1_ZA_SMP_CD:           xlApp.Range("E8").Value = ss1.Text
    ss1.Col = SS1_ZXB:           xlApp.Range("F8").Value = ss1.Text
    ss1.Col = SS1_ZB_SMP_CD:           xlApp.Range("G8").Value = ss1.Text
    ss1.Col = SS1_ZXC:           xlApp.Range("H8").Value = ss1.Text
    ss1.Col = SS1_ZC_SMP_CD:           xlApp.Range("I8").Value = ss1.Text
    ss1.Col = SS1_ZXD:           xlApp.Range("J8").Value = ss1.Text
    ss1.Col = SS1_ZD_SMP_CD:           xlApp.Range("K8").Value = ss1.Text
    ss1.Col = SS1_ZXE:           xlApp.Range("L8").Value = ss1.Text
    ss1.Col = SS1_ZE_SMP_CD:           xlApp.Range("M8").Value = ss1.Text
    ss1.Col = SS1_ZXF:           xlApp.Range("N8").Value = ss1.Text
    ss1.Col = SS1_ZF_SMP_CD:           xlApp.Range("O8").Value = ss1.Text
    
    ss1.Col = SS1_WQA:           xlApp.Range("P8").Value = ss1.Text
    ss1.Col = SS1_WQ_SMP_CD:           xlApp.Range("Q8").Value = ss1.Text
    ss1.Col = SS1_YDA:           xlApp.Range("R8").Value = ss1.Text
    ss1.Col = SS1_YD_SMP_CD:           xlApp.Range("S8").Value = ss1.Text
    ss1.Col = SS1_JXA:           xlApp.Range("T8").Value = ss1.Text
    ss1.Col = SS1_XA_SMP_CD:           xlApp.Range("U8").Value = ss1.Text
    ss1.Col = SS1_JZA:           xlApp.Range("V8").Value = ss1.Text
    ss1.Col = SS1_JA_SMP_CD:           xlApp.Range("W8").Value = ss1.Text
    
    ss1.Col = SS1_TEST:          xlApp.Range("A9").Value = "备注：" & ss1.Text
    
    
'    ss1.Col = SS1_TEST:          sText = ss1.Text
'    ss1.Col = SS1_SMP_CD:        sSmpcd = ss1.Text
'
''    ss1.Col = SS1_TEST:          xlApp.Range("A9").Value = "备注：" & ss1.Text & " ； " &
'    xlApp.Range("A9").Value = "备注：" & sText & " ； " & sSmpcd
    
    If ss1.MaxRows > 1 Then
        ss1.Row = 2
        ss1.Col = SS1_SMPNO:         xlApp.Range("A11").Value = ss1.Text
        ss1.Col = SS1_STDSPEC:       xlApp.Range("B11").Value = ss1.Text
    '    ss1.Col = SS1_SIZE:          xlApp.Range("C11").Value = ss1.Text
    '    xlApp.Range("C11").Value = ss1.Col = SS1_SIZE & ",数量：" & ss1.Col = SS1_NO
'        xlApp.Range("C11").Value = TXT_SIZE.Text & ",数量：" & TXT_NO.Text
        With ss1
            .Row = 2:   .Col = SS1_SIZE
            sSize = .Text
            .Row = 2:   .Col = SS1_NO
            sNo = .Text
            xlApp.Range("C11").Value = sSize & ",数量：" & sNo
        End With
        ss1.Col = SS1_LSA:            xlApp.Range("D11").Value = ss1.Text
        ss1.Col = SS1_LA_SMP_CD:           xlApp.Range("E11").Value = ss1.Text
        ss1.Col = SS1_LSB:            xlApp.Range("F11").Value = ss1.Text
        ss1.Col = SS1_LB_SMP_CD:           xlApp.Range("G11").Value = ss1.Text
        ss1.Col = SS1_LSC:            xlApp.Range("H11").Value = ss1.Text
        ss1.Col = SS1_LC_SMP_CD:           xlApp.Range("I11").Value = ss1.Text
        ss1.Col = SS1_LSD:            xlApp.Range("J11").Value = ss1.Text
        ss1.Col = SS1_LD_SMP_CD:           xlApp.Range("K11").Value = ss1.Text
        ss1.Col = SS1_LSE:           xlApp.Range("L11").Value = ss1.Text
        ss1.Col = SS1_LE_SMP_CD:           xlApp.Range("M11").Value = ss1.Text
        ss1.Col = SS1_LSF:           xlApp.Range("N11").Value = ss1.Text
        ss1.Col = SS1_LF_SMP_CD:           xlApp.Range("O11").Value = ss1.Text
        ss1.Col = SS1_LSG:           xlApp.Range("P11").Value = ss1.Text
        ss1.Col = SS1_LG_SMP_CD:           xlApp.Range("Q11").Value = ss1.Text
        ss1.Col = SS1_LSH:           xlApp.Range("R11").Value = ss1.Text
        ss1.Col = SS1_LH_SMP_CD:           xlApp.Range("S11").Value = ss1.Text
        ss1.Col = SS1_LSI:           xlApp.Range("T11").Value = ss1.Text
        ss1.Col = SS1_LI_SMP_CD:           xlApp.Range("U11").Value = ss1.Text
        ss1.Col = SS1_LSJ:           xlApp.Range("V11").Value = ss1.Text
        ss1.Col = SS1_LJ_SMP_CD:           xlApp.Range("W11").Value = ss1.Text
        ss1.Col = SS1_LSK:           xlApp.Range("X11").Value = ss1.Text
        ss1.Col = SS1_LK_SMP_CD:           xlApp.Range("Y11").Value = ss1.Text
        ss1.Col = SS1_LSL:           xlApp.Range("Z11").Value = ss1.Text
        ss1.Col = SS1_LL_SMP_CD:           xlApp.Range("AA11").Value = ss1.Text
        
        ss1.Col = SS1_CJA:           xlApp.Range("D13").Value = ss1.Text
        ss1.Col = SS1_CA_SMP_CD:           xlApp.Range("E13").Value = ss1.Text
        ss1.Col = SS1_CJB:           xlApp.Range("F13").Value = ss1.Text
        ss1.Col = SS1_CB_SMP_CD:           xlApp.Range("G13").Value = ss1.Text
        ss1.Col = SS1_CJC:           xlApp.Range("H13").Value = ss1.Text
        ss1.Col = SS1_CC_SMP_CD:           xlApp.Range("I13").Value = ss1.Text
        ss1.Col = SS1_CJD:           xlApp.Range("J13").Value = ss1.Text
        ss1.Col = SS1_CD_SMP_CD:           xlApp.Range("K13").Value = ss1.Text
        ss1.Col = SS1_CJE:           xlApp.Range("L13").Value = ss1.Text
        ss1.Col = SS1_CE_SMP_CD:           xlApp.Range("M13").Value = ss1.Text
        ss1.Col = SS1_CJF:           xlApp.Range("N13").Value = ss1.Text
        ss1.Col = SS1_CF_SMP_CD:           xlApp.Range("O13").Value = ss1.Text
        ss1.Col = SS1_CJG:           xlApp.Range("P13").Value = ss1.Text
        ss1.Col = SS1_CG_SMP_CD:           xlApp.Range("Q13").Value = ss1.Text
        ss1.Col = SS1_CJH:           xlApp.Range("R13").Value = ss1.Text
        ss1.Col = SS1_CH_SMP_CD:           xlApp.Range("S13").Value = ss1.Text
        ss1.Col = SS1_CJI:           xlApp.Range("T13").Value = ss1.Text
        ss1.Col = SS1_CI_SMP_CD:           xlApp.Range("U13").Value = ss1.Text
        ss1.Col = SS1_CJJ:           xlApp.Range("V13").Value = ss1.Text
        ss1.Col = SS1_CJ_SMP_CD:           xlApp.Range("W13").Value = ss1.Text
        ss1.Col = SS1_CJK:           xlApp.Range("X13").Value = ss1.Text
        ss1.Col = SS1_CK_SMP_CD:           xlApp.Range("Y13").Value = ss1.Text
        ss1.Col = SS1_CJL:           xlApp.Range("Z13").Value = ss1.Text
        ss1.Col = SS1_CL_SMP_CD:           xlApp.Range("AA13").Value = ss1.Text
        
        ss1.Col = SS1_ZXA:           xlApp.Range("D15").Value = ss1.Text
        ss1.Col = SS1_ZA_SMP_CD:           xlApp.Range("E15").Value = ss1.Text
        ss1.Col = SS1_ZXB:           xlApp.Range("F15").Value = ss1.Text
        ss1.Col = SS1_ZB_SMP_CD:           xlApp.Range("G15").Value = ss1.Text
        ss1.Col = SS1_ZXC:           xlApp.Range("H15").Value = ss1.Text
        ss1.Col = SS1_ZC_SMP_CD:           xlApp.Range("I15").Value = ss1.Text
        ss1.Col = SS1_ZXD:           xlApp.Range("J15").Value = ss1.Text
        ss1.Col = SS1_ZD_SMP_CD:           xlApp.Range("K15").Value = ss1.Text
        ss1.Col = SS1_ZXE:           xlApp.Range("L15").Value = ss1.Text
        ss1.Col = SS1_ZE_SMP_CD:           xlApp.Range("M15").Value = ss1.Text
        ss1.Col = SS1_ZXF:           xlApp.Range("N15").Value = ss1.Text
        ss1.Col = SS1_ZF_SMP_CD:           xlApp.Range("O15").Value = ss1.Text
        
        ss1.Col = SS1_WQA:           xlApp.Range("P15").Value = ss1.Text
        ss1.Col = SS1_WQ_SMP_CD:           xlApp.Range("Q15").Value = ss1.Text
        ss1.Col = SS1_YDA:           xlApp.Range("R15").Value = ss1.Text
        ss1.Col = SS1_YD_SMP_CD:           xlApp.Range("S15").Value = ss1.Text
        ss1.Col = SS1_JXA:           xlApp.Range("T15").Value = ss1.Text
        ss1.Col = SS1_XA_SMP_CD:           xlApp.Range("U15").Value = ss1.Text
        ss1.Col = SS1_JZA:           xlApp.Range("V15").Value = ss1.Text
        ss1.Col = SS1_JA_SMP_CD:           xlApp.Range("W15").Value = ss1.Text
    
        ss1.Col = SS1_TEST:          xlApp.Range("A16").Value = "备注：" & ss1.Text
    
    End If
    
    
'    Clipboard.Clear
'    ss1.SetSelection 1, 1, 8, ss1.MaxRows
'    ss1.ClipboardCopy
'    xlApp.Range("A7").Select
'    xlApp.ActiveSheet.Paste
'    Clipboard.Clear

    
    ss1.ClearSelection
       
    Screen.MousePointer = vbDefault
    
    xlApp.Application.Visible = True
    
    xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
    
    Set xlSheet = Nothing
    Set xlApp = Nothing
        
    Exit Sub

ErrHandle:
    MsgBox Error
    
    Set xlSheet = Nothing
    Set xlApp = Nothing
    Screen.MousePointer = vbDefault
End Sub

