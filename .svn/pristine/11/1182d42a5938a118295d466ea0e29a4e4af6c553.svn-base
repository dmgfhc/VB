VERSION 5.00
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form AGC2432C 
   Caption         =   "理化检验委托单_AGC2432C"
   ClientHeight    =   8145
   ClientLeft      =   180
   ClientTop       =   2985
   ClientWidth     =   15735
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8145
   ScaleWidth      =   15735
   WindowState     =   2  'Maximized
   Begin VB.CheckBox txt_check 
      BackColor       =   &H00E0E0E0&
      Caption         =   "已处理对象"
      Height          =   255
      Left            =   5190
      TabIndex        =   13
      Top             =   80
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
      Top             =   80
      Width           =   735
   End
   Begin VB.ComboBox txt_HTM_METH 
      Height          =   300
      ItemData        =   "AGC2432C.frx":0000
      Left            =   13170
      List            =   "AGC2432C.frx":000D
      TabIndex        =   2
      Top             =   80
      Width           =   1065
   End
   Begin VB.CheckBox Chc_OutOrder 
      BackColor       =   &H00E0E0E0&
      Caption         =   "可委托出口船板"
      Height          =   255
      Left            =   6510
      TabIndex        =   1
      Top             =   80
      Width           =   1695
   End
   Begin VB.TextBox Txt_OutOrder 
      Height          =   270
      Left            =   11310
      TabIndex        =   0
      Text            =   "0"
      Top             =   80
      Visible         =   0   'False
      Width           =   375
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8295
      Left            =   0
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   855
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
      MaxCols         =   35
      MaxRows         =   1
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      SpreadDesigner  =   "AGC2432C.frx":001A
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   30
      Top             =   80
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
         Size            =   9.76
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
      Top             =   80
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
         Size            =   9.76
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
      Top             =   80
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
      Top             =   80
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
      Top             =   80
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
      Top             =   80
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
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
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

Const SS1_REQ_END_DATE = 27
Const SS1_SMP_CUT_PLT = 32
Const SS1_SMP_NO = 28
Const SS1_URGNT_FL = 34  '紧急订单绿色标记 2012-11-08  by  LiQian


Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"
         
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
     Call Gp_Ms_Collection(DTP_PROD_FR, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(DTP_PROD_TO, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'    Call Gp_Ms_Collection(TXT_SLAB_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_Plt, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_LINE, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
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
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 35, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AGC2430C.P_ONEROW", Key:="P-O"
    sc1.Add Item:="AGC2430C.P_REFER", Key:="P-R"
    sc1.Add Item:="AGC2430C.P_MODIFY", Key:="P-M"
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
    Call Gp_Sp_ColHidden(ss1, 33, True)
    Call Gp_Sp_ColHidden(ss1, 30, True)
    Call Gp_Sp_ColHidden(ss1, 31, True)
    Call Gp_Sp_ColHidden(ss1, 32, True)
    
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
    
    
'    Set AdoRs = New adodb.Recordset
'
'    sQuery = "SELECT TO_CHAR(SYSDATE,'YYYY-MM-DD HH24:MI:SS') FROM DUAL"
'    AdoRs.Open sQuery, M_CN1, adOpenKeyset
'
'    sDate = AdoRs.Fields(0)
'
'    AdoRs.Close
'    Set AdoRs = Nothing

'   增加实验条件 2012.5.29 LIUXIANG
   
    Set AdoRs = New ADODB.Recordset
    
    sQuery = "SELECT    '',A.SMP_NO,A.STDSPEC ,A.THK  ,A.SMP_CNT ,A.TENCIL_FL  ,A.HGT_TENCIL_FL, A.Bend_Fl,A.Impact_TEMP  ,A.Drop_Wgt_TEMP"
    sQuery = sQuery + " ,A.Tim_Imact_TEMP    ,DECODE(A.Non_Metal_Fl,'1','Y','') ,DECODE(A.MACRO_FL,1,'Y','') ,A.Hardness_Fl,DECODE(A.Chem_Fl,1,'Y','')   ,DECODE(A.Ton_Fl,1,'Y','')   ,DECODE(A.Std_Smp_Fl,1,'Y','')"
    sQuery = sQuery + " ,A.Photo_Fl  ,SUBSTR(A.Text,1,48), A.WRK_DATE     ,A.UPD_EMP      ,gf_empnamefind(A.UPD_EMP)   ,A.UPD_DATE"
    sQuery = sQuery + " ,A.UPD_TIME  ,DECODE (B.PRC,'DH','热处理'||B.PRC_LINE , GF_COMNNAMEFIND('C0001',B.SMP_CUT_PLT) )   ,'1'"
    sQuery = sQuery + " ,GF_COMNNAMEFIND('Q0089',C.YP_TYPE_CD), GF_COMNNAMEFIND('Q0089',C.A_YP_TYPE_CD), GF_COMNNAMEFIND('Q0089',C.HGT_YP_TYPE_CD), GF_COMNNAMEFIND('Q0089',C.A_HGT_YP_TYPE_CD)"
    sQuery = sQuery + " ,GF_COMNNAMEFIND('Q0008',C.IMPACT_KND), GF_COMNNAMEFIND('Q0008',C.A_IMPACT_KND), GF_COMNNAMEFIND('Q0008',C.TIM_IMPACT_KND), GF_COMNNAMEFIND('Q0008',C.A_TIM_IMPACT_KND)"
    sQuery = sQuery + " ,GF_COMNNAMEFIND('Q0057',C.IMPACT_SIZE_CD), GF_COMNNAMEFIND('Q0057',C.A_IMPACT_SIZE_CD), GF_COMNNAMEFIND('Q0057',C.TIM_IMPACT_SIZE_CD), GF_COMNNAMEFIND('Q0057',C.A_TIM_IMPACT_SIZE_CD)"
    sQuery = sQuery + " , DECODE(A.HIC_Fl,1,'Y',''), GF_COMNNAMEFIND('Q0090',C.HIC_STD_CD), C.HIC_SVT_KND"
    sQuery = sQuery + " FROM   Qp_Smp_Send A,QP_TEST_HEAD B,QP_QLTY_MATR C"
    sQuery = sQuery + " WHERE  A.SMP_NO      = B.SMP_NO"
    sQuery = sQuery + "   AND  A.SMP_SEND_NO = '" & txt_smp_sent_no & "'"
    sQuery = sQuery + "   AND  B.ORD_NO = C.ORD_NO"
    sQuery = sQuery + "   AND  B.ORD_ITEM = C.ORD_ITEM"
    sQuery = sQuery + "   AND C.KND = (SELECT  MAX(KND) FROM  NISCO.QP_QLTY_MATR "
    sQuery = sQuery + " WHERE ORD_NO = B.ORD_NO AND ORD_ITEM = B.ORD_ITEM AND KND IN('1','2'))"
    sQuery = sQuery + " ORDER BY A.SMP_SEQ,A.STDSPEC,A.THK ,A.SMP_NO"
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    If AdoRs.EOF Then
        AdoRs.Close
        Exit Sub
    End If
    
    arrRecords1 = AdoRs.GetRows
    AdoRs.Close

    Set AdoRs = Nothing
       
    Call SAMPLE_SEND_PRINT(arrRecords1)
    
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

    sQuery = "{call AGC2430C.P_MODIFY2 ( '" + txt_smp_sent_no.Text + "',?,? )}"
    
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

    sQuery = "{call AQC1060P ('" + txt_smp_sent_no.Text + "','',?,?)}"
    
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
       txt_Plt = "C1"
    ElseIf App.Title = "CG" Then
       txt_Plt = "C3"
    End If
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "Q-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault
    
    txt_SMP_FL.Value = 0
    txt_check.Value = 0
    Call subButtonHide
    txt_smp_sent_no.Text = ""
    TXT_CUT_PLT.Text = ""
    
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
            
            '紧急订单绿色显示 add by liqian 2012-11-08
             With ss1
                  For iRow = 1 To .MaxRows
                     .Row = iRow:
                      .Col = SS1_URGNT_FL:
                      sUrgnt_Fl = Trim(.Text)
                      .Col = 35:                '可委托出口船板
                      If (Trim(.Text) > 0) Then
                       Call Gp_Sp_BlockColor(ss1, 1, 1, iRow, iRow, vbBlack, &HFF80FF)
                      End If
                                            
                      If sUrgnt_Fl = "Y" Then
                         Call Gp_Sp_BlockColor(ss1, 1, .MaxCols, iRow, iRow, &HC000&)
                      End If
                  Next iRow
            End With
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
    Dim I           As Integer
    Dim sREQ_DATE   As String
    Dim COUNT1   As Integer

    COUNT1 = 0

    If txt_check.Value = "0" Then
    
        Set AdoRs = New ADODB.Recordset
           
        sQuery = "SELECT Gf_SMP_SEND_NO( "
        sQuery = sQuery & "'" & txt_Plt & "') "
        sQuery = sQuery & "FROM DUAL"
        
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
        If Not AdoRs.BOF And Not AdoRs.EOF Then
           For I = 1 To ss1.MaxRows
               ss1.Row = I
               
               ss1.Col = 0
               If ss1.Text = "Update" Or ss1.Text = "Insert" Then
               
                  ss1.Col = 2
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
                 If COUNT1 > 24 Then
                    Call MsgBox("一张委托单不能超过24个试样！", vbCritical, "系统提示信息")
                    Exit Sub
                 End If
              End If
           
           Next I
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

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)

    If txt_check.Value = 1 Then

        ss1.Row = ss1.ActiveRow
        ss1.Col = 2
        txt_smp_sent_no.Text = ss1.Text
        ss1.Col = SS1_SMP_CUT_PLT
        TXT_CUT_PLT.Text = ss1.Text
    End If
    
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), SS1_SMP_NO)
    End If
    
'HYS INSERT START
    Dim sREQ_DATE As String
    Dim sUrgnt_Fl As String
    Dim OutOrder As String
        ss1.Row = Row
        ss1.Col = 1
        
        If ss1.Text = "1" Then
'           COUNT1 = COUNT1 + 1
'           If COUNT1 > 36 Then
'              Call MsgBox("一张委托单不能超过36个试样！", vbCritical, "系统提示信息")
'              ss1.Text = "0"
'              Exit Sub
'           End If
           
           ss1.Col = SS1_REQ_END_DATE
           sREQ_DATE = ss1.Text
           If sREQ_DATE = "" Then
              ss1.Text = dtp_end_date.RawData
           End If
        Else
           Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
             ss1.Col = 35:
             ss1.Row = Row
             If Trim(ss1.Text) > 0 Then
                Call Gp_Sp_BlockColor(ss1, 1, 1, Row, Row, vbBlack, &HFF80FF)
             End If
        End If
'HYS INSERT END
    
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

        Call Gp_Sp_ColHidden(ss1, 2, False)
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

        Call Gp_Sp_ColHidden(ss1, 2, True)
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

Private Function SAMPLE_SEND_PRINT(arrRecords1 As Variant) As String
    Dim RowCnt      As Long
    Dim PrtCnt      As Long
    Dim LneCnt      As Long
    Dim sRow        As String
    Dim ROW_NUM     As Integer
    Dim pAry()      As String
    

    If TXT_CUT_PLT.Text = "" Then
       If txt_DH_FL.Value = "1" Then
          TXT_CUT_PLT.Text = "热处理" & txt_LINE.Text
       Else
          TXT_CUT_PLT.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_Plt.Text), 1)
       End If
    End If
    
    If IsEmpty(arrRecords1) Then
        SAMPLE_SEND_PRINT = "Err Data"
        Exit Function
    End If
        
    RowCnt = UBound(arrRecords1, 2)
    
    PrtCnt = -1
    LneCnt = 0
    ROW_NUM = 0
    
    ReDim pAry(1 To 24, 1 To 21)

    
    Do
 
        LneCnt = LneCnt + 1
        PrtCnt = PrtCnt + 1

        pAry(LneCnt, 1) = ROW_NUM + LneCnt                               ' SEQ
        pAry(LneCnt, 2) = arrRecords1(1, PrtCnt) & ""                    ' SMP_NO
        pAry(LneCnt, 3) = arrRecords1(2, PrtCnt) & ""                    ' STDSPEC
        pAry(LneCnt, 4) = arrRecords1(3, PrtCnt) & ""                    ' THK
        pAry(LneCnt, 5) = arrRecords1(4, PrtCnt) & ""                    ' SMP_CNT
        
        pAry(LneCnt, 6) = arrRecords1(5, PrtCnt) & ""                    ' TENCIL_FL
        pAry(LneCnt, 7) = arrRecords1(6, PrtCnt) & ""                    ' TENCIL_FL
        pAry(LneCnt, 8) = arrRecords1(7, PrtCnt) & ""                    ' Bend_Fl

        pAry(LneCnt, 9) = arrRecords1(8, PrtCnt) & ""                    ' Impact_TEMP
        pAry(LneCnt, 10) = arrRecords1(9, PrtCnt) & ""                    ' Drop_Wgt_TEMP
        pAry(LneCnt, 11) = arrRecords1(10, PrtCnt) & ""                   ' Tim_Imact_TEMP
        
        pAry(LneCnt, 12) = arrRecords1(11, PrtCnt) & ""                  ' MACRO_FL
        pAry(LneCnt, 13) = arrRecords1(12, PrtCnt) & ""                  ' Non_Metal_Fl
        pAry(LneCnt, 14) = arrRecords1(13, PrtCnt) & ""                  ' Hardness_Fl
        pAry(LneCnt, 15) = arrRecords1(14, PrtCnt) & ""                  ' Chem_Fl
        pAry(LneCnt, 16) = arrRecords1(15, PrtCnt) & ""                  ' Ton_Fl
        pAry(LneCnt, 17) = arrRecords1(16, PrtCnt) & ""                  ' Std_Smp_Fl
        pAry(LneCnt, 18) = arrRecords1(17, PrtCnt) & ""                  ' Photo_Fl
        pAry(LneCnt, 19) = arrRecords1(38, PrtCnt) & ""                  ' HIC_Fl
        pAry(LneCnt, 20) = arrRecords1(18, PrtCnt) & ""                  ' Text
        
        
        
        '   增加实验条件 2012.5.29 LIUXIANG
        If Not IsNull(arrRecords1(26, PrtCnt)) Then
            pAry(LneCnt, 21) = pAry(LneCnt, 21) + "屈服类型：" + arrRecords1(26, PrtCnt) & " "
        End If
        If Not IsNull(arrRecords1(27, PrtCnt)) Then
            pAry(LneCnt, 21) = pAry(LneCnt, 21) + "追加屈服类型：" + arrRecords1(27, PrtCnt) & " "
        End If
        If Not IsNull(arrRecords1(28, PrtCnt)) Then
            pAry(LneCnt, 21) = pAry(LneCnt, 21) + "高温屈服类型：" + arrRecords1(28, PrtCnt) & " "
        End If
        If Not IsNull(arrRecords1(29, PrtCnt)) Then
            pAry(LneCnt, 21) = pAry(LneCnt, 21) + "追加高温屈服类型：" + arrRecords1(29, PrtCnt) & " "
        End If
        
        If Not IsNull(arrRecords1(30, PrtCnt)) Then
            pAry(LneCnt, 21) = pAry(LneCnt, 21) + "冲击开槽：" + arrRecords1(30, PrtCnt) & " "
        End If
        If Not IsNull(arrRecords1(31, PrtCnt)) Then
            pAry(LneCnt, 21) = pAry(LneCnt, 21) + "追加冲击开槽：" + arrRecords1(31, PrtCnt) & " "
        End If
        If Not IsNull(arrRecords1(32, PrtCnt)) Then
            pAry(LneCnt, 21) = pAry(LneCnt, 21) + "时效冲击开槽：" + arrRecords1(32, PrtCnt) & " "
        End If
        If Not IsNull(arrRecords1(33, PrtCnt)) Then
            pAry(LneCnt, 21) = pAry(LneCnt, 21) + "追加时效冲击开槽：" + arrRecords1(33, PrtCnt) & " "
        End If
         
        
        If Not IsNull(arrRecords1(34, PrtCnt)) Then
            pAry(LneCnt, 21) = pAry(LneCnt, 21) + "冲击尺寸：" + arrRecords1(34, PrtCnt) & " "
        End If
        If Not IsNull(arrRecords1(35, PrtCnt)) Then
            pAry(LneCnt, 21) = pAry(LneCnt, 21) + "追加冲击尺寸：" + arrRecords1(35, PrtCnt) & " "
        End If
        If Not IsNull(arrRecords1(36, PrtCnt)) Then
            pAry(LneCnt, 21) = pAry(LneCnt, 21) + "时效冲击尺寸：" + arrRecords1(36, PrtCnt) & " "
        End If
        If Not IsNull(arrRecords1(37, PrtCnt)) Then
            pAry(LneCnt, 21) = pAry(LneCnt, 21) + "追加时效冲击尺寸：" + arrRecords1(37, PrtCnt) & " "
        End If
        
        If Not IsNull(arrRecords1(39, PrtCnt)) Then
            pAry(LneCnt, 21) = pAry(LneCnt, 21) + "HIC试验标准：" + arrRecords1(39, PrtCnt) & " "
        End If
        If Not IsNull(arrRecords1(40, PrtCnt)) Then
            pAry(LneCnt, 21) = pAry(LneCnt, 21) + "溶液类型：" + arrRecords1(40, PrtCnt) & " "
        End If


       
        If LneCnt = 24 Then
            ROW_NUM = LneCnt + ROW_NUM
            Set xlApp = GetObject("", "Excel.Application")
            If Err.Number = 429 Then
                Set xlApp = CreateObject("", "Excel.Application")
            End If
            
            xlApp.Workbooks.Open (App.Path & "\AGC2430C.xls")
            Set xlSheet = xlApp.Worksheets("Sheet1")
            
            xlApp.Range("A3").Value = "委托单号：" & txt_smp_sent_no.Text
            xlApp.Range("A4").Value = "委托单位：" & TXT_CUT_PLT.Text
            xlApp.Range("D4").Value = Mid(Now, 1, 4) + "年" + _
                                      Mid(Now, 6, 2) + "月" + _
                                      Mid(Now, 9, 2) + "日"
            
            
            sRow = "A7" & ":U" & 6 + LneCnt
            xlSheet.Range(sRow).Value = pAry
            
            xlApp.Range(sRow).Select
            With xlApp.Selection.Borders
                .LineStyle = 1
            End With
            
            '
            
    
            sRow = "B" & LneCnt + 7
            xlApp.Range(sRow).Value = "委托人：" & sUserName
            xlApp.Range(sRow).Font.Size = 10
            sRow = "E" & LneCnt + 7
            xlApp.Range(sRow).Value = "委托时间：" & Format(Now, "YYYY-MM-DD HH:MM:SS")
            xlApp.Range(sRow).Font.Size = 10
            sRow = "J" & LneCnt + 7
            xlApp.Range(sRow).Value = "送样人："
            xlApp.Range(sRow).Font.Size = 10
            sRow = "O" & LneCnt + 7
            xlApp.Range(sRow).Value = "送样时间："
            xlApp.Range(sRow).Font.Size = 10
            
'            xlApp.ActiveSheet.Paste
    
            Screen.MousePointer = vbDefault
'            Set xlSheet = Nothing
'            Set xlApp = Nothing
            xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
'            xlApp.Application.Visible = True
'            xlApp.ActiveWindow.SelectedSheets.PrintPreview
 
            Set xlSheet = Nothing
            xlApp.ActiveWorkbook.Close False
            xlApp.Quit

            LneCnt = 0
            
            ReDim pAry(1 To 24, 1 To 21)
            
        End If

    Loop Until PrtCnt = RowCnt
    
'--------------------------------------------------------------------------------------------------------
    If LneCnt <> 0 Then
    
        
        Set xlApp = GetObject("", "Excel.Application")
        If Err.Number = 429 Then
            Set xlApp = CreateObject("", "Excel.Application")
        End If
        
        xlApp.Workbooks.Open (App.Path & "\AGC2430C.xls")
        Set xlSheet = xlApp.Worksheets("Sheet1")
        
        xlApp.Range("A3").Value = "委托单号：" & txt_smp_sent_no.Text
        xlApp.Range("A4").Value = "委托单位：" & TXT_CUT_PLT.Text
        xlApp.Range("D4").Value = Mid(Now, 1, 4) + "年" + _
                                  Mid(Now, 6, 2) + "月" + _
                                  Mid(Now, 9, 2) + "日"
         
        sRow = "A7" & ":U" & 6 + LneCnt
        xlSheet.Range(sRow).Value = pAry
        
        xlApp.Range(sRow).Select
        With xlApp.Selection.Borders
            .LineStyle = 1
        End With

        sRow = "B" & LneCnt + 7
        xlApp.Range(sRow).Value = "委托人：" & sUserName
        xlApp.Range(sRow).Font.Size = 10
        sRow = "E" & LneCnt + 7
        xlApp.Range(sRow).Value = "委托时间：" & Format(Now, "YYYY-MM-DD HH:MM:SS")
        xlApp.Range(sRow).Font.Size = 10
        sRow = "J" & LneCnt + 7
        xlApp.Range(sRow).Value = "送样人："
        xlApp.Range(sRow).Font.Size = 10
        sRow = "O" & LneCnt + 7
        xlApp.Range(sRow).Value = "送样时间："
        xlApp.Range(sRow).Font.Size = 10
        
'            xlApp.ActiveSheet.Paste
        
        Screen.MousePointer = vbDefault
        
'        Set xlSheet = Nothing
'        Set xlApp = Nothing
         xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
'         xlApp.Application.Visible = True
'         xlApp.ActiveWindow.SelectedSheets.PrintPreview
         
        Set xlSheet = Nothing
        xlApp.ActiveWorkbook.Close False
        xlApp.Quit

    End If
    
    Set xlApp = Nothing
    
    Exit Function
    
End Function




