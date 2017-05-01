VERSION 5.00
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form DGC2430C 
   Caption         =   "理化检验委托单_DGC2430C"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.ComboBox txt_plt 
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
      ItemData        =   "DGC2430C.frx":0000
      Left            =   1320
      List            =   "DGC2430C.frx":000A
      TabIndex        =   14
      Tag             =   "工厂"
      Top             =   555
      Width           =   735
   End
   Begin VB.ComboBox COB_GROUP 
      Height          =   300
      ItemData        =   "DGC2430C.frx":0016
      Left            =   9510
      List            =   "DGC2430C.frx":0026
      TabIndex        =   9
      Top             =   30
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.ComboBox COB_SHIFT 
      Height          =   300
      ItemData        =   "DGC2430C.frx":0036
      Left            =   11880
      List            =   "DGC2430C.frx":0043
      TabIndex        =   8
      Top             =   30
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.CheckBox txt_check 
      Caption         =   "已处理对象"
      Height          =   390
      Left            =   6870
      TabIndex        =   7
      Top             =   75
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   525
      Left            =   2190
      TabIndex        =   4
      Top             =   405
      Width           =   2400
      Begin VB.CheckBox txt_DH_FL 
         Caption         =   "热处理"
         Height          =   240
         Left            =   75
         TabIndex        =   6
         Top             =   195
         Width           =   915
      End
      Begin VB.TextBox txt_line 
         Height          =   300
         Left            =   1785
         MaxLength       =   10
         TabIndex        =   5
         Top             =   135
         Width           =   450
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Left            =   1110
         Top             =   135
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
      Height          =   300
      Left            =   14190
      MaxLength       =   13
      TabIndex        =   3
      Top             =   60
      Width           =   1320
   End
   Begin VB.TextBox TXT_CUT_PLT 
      Height          =   330
      Left            =   4845
      TabIndex        =   2
      Top             =   60
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
      Height          =   480
      Left            =   10785
      TabIndex        =   1
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
      Height          =   480
      Left            =   12975
      TabIndex        =   0
      Top             =   360
      Width           =   1725
   End
   Begin InDate.UDate dtp_prod_fr 
      Height          =   315
      Left            =   1320
      TabIndex        =   10
      Tag             =   "日期"
      Top             =   90
      Width           =   1455
      _ExtentX        =   2566
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
      Left            =   10650
      Top             =   30
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "班次"
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
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   8280
      Top             =   30
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "班别"
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
   Begin InDate.UDate dtp_prod_to 
      Height          =   315
      Left            =   3210
      TabIndex        =   11
      Top             =   90
      Width           =   1455
      _ExtentX        =   2566
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
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   2805
      Top             =   90
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   556
      Caption         =   "至"
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
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   12840
      Top             =   60
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
      Height          =   300
      Left            =   6870
      Top             =   510
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
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
      ForeColor       =   16711680
   End
   Begin InDate.UDate dtp_end_date 
      Height          =   315
      Left            =   8235
      TabIndex        =   12
      Top             =   510
      Width           =   1455
      _ExtentX        =   2566
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
   Begin FPSpread.vaSpread ss1 
      Height          =   8370
      Left            =   85
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1020
      Width           =   15000
      _Version        =   393216
      _ExtentX        =   26458
      _ExtentY        =   14764
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
      MaxCols         =   28
      MaxRows         =   1
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      SpreadDesigner  =   "DGC2430C.frx":0050
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   85
      Top             =   90
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
      Left            =   85
      Top             =   555
      Width           =   1200
      _ExtentX        =   2117
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
      ForeColor       =   16711680
   End
End
Attribute VB_Name = "DGC2430C"
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
Dim COUNT1   As Integer
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

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"
         
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
     Call Gp_Ms_Collection(dtp_prod_fr, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(dtp_prod_to, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'    Call Gp_Ms_Collection(TXT_SLAB_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_plt, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_line, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(txt_smp_sent_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_check, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_DH_FL, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
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
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
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
    Call Gp_Sp_ColHidden(ss1, 23, True)
    Call Gp_Sp_ColHidden(ss1, 28, True)
    Call Gp_Sp_ColHidden(ss1, 25, True)
    Call Gp_Sp_ColHidden(ss1, 26, True)
    Call Gp_Sp_ColHidden(ss1, 27, True)
    
    Me.KeyPreview = True

End Sub



Private Sub cmdReport_Click()

    Dim sQuery As String
    Dim arrRecords1 As Variant
    Dim AdoRs As ADODB.Recordset
    
    If ss1.MaxRows < 1 Then Exit Sub
    
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
    

   
    Set AdoRs = New ADODB.Recordset
    
    sQuery = "SELECT    '',A.SMP_NO,A.STDSPEC ,A.THK  ,A.SMP_CNT ,DECODE(A.TENCIL_FL,1,'Y','')    , DECODE(A.Bend_Fl,1,'Y','')  ,A.Impact_TEMP  ,A.Tim_Imact_TEMP"
    sQuery = sQuery + " ,A.Drop_Wgt_TEMP    ,DECODE(A.MACRO_FL,1,'Y','')  ,DECODE(A.Non_Metal_Fl,1,'Y','') ,DECODE(A.Hardness_Fl,1,'Y','')  ,DECODE(A.Chem_Fl,1,'Y','')   ,DECODE(A.Ton_Fl,1,'Y','')   ,DECODE(A.Std_Smp_Fl,1,'Y','')"
    sQuery = sQuery + " ,DECODE(A.Photo_Fl,1,'Y','')  ,substr(A.Text,1,5) ,A.WRK_DATE     ,A.UPD_EMP      ,gf_empnamefind(A.UPD_EMP)   ,A.UPD_DATE"
    sQuery = sQuery + " ,A.UPD_TIME  ,DECODE (B.PRC,'DH','热处理'||B.PRC_LINE , GF_COMNNAMEFIND('C0001',B.SMP_CUT_PLT) )   ,'1'"
    sQuery = sQuery + " FROM   Qp_Smp_Send A,QP_TEST_HEAD B"
    sQuery = sQuery + " WHERE  A.SMP_NO = B.SMP_NO"
    sQuery = sQuery + "   AND  A.SMP_SEND_NO = '" & txt_smp_sent_no & "'"
    sQuery = sQuery + " ORDER BY A.STDSPEC,A.THK ,A.SMP_NO"
    
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
    
    dtp_prod_fr.RawData = Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYYMMDD') FROM DUAL")
    dtp_prod_to.RawData = Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYYMMDD') FROM DUAL")

'    If App.Title = "BG" Then
       txt_plt = "C1"
'    ElseIf App.Title = "CG" Then
'       TXT_PLT = "C3"
'    Else
'
'    End If
    
    Screen.MousePointer = vbDefault
    
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

    Dim sMesg As String
    
    COUNT1 = 0
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
                
        If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
            ss1.OperationMode = OperationModeNormal
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call subButtonHide
            
            If ss1.MaxRows >= 1 Then
               ss1.Row = 1
               ss1.Col = 27
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
    
    If txt_check.Value = "0" Then
    
        Set AdoRs = New ADODB.Recordset
           
        sQuery = "SELECT Gf_SMP_SEND_NO( "
        sQuery = sQuery & "'" & txt_plt & "') "
        sQuery = sQuery & "FROM DUAL"
        
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
        If Not AdoRs.BOF And Not AdoRs.EOF Then
           For I = 1 To ss1.MaxRows
               ss1.Row = I
               ss1.Col = 0
               If ss1.Text = "Update" Or ss1.Text = "Insert" Then
               
                  ss1.Col = 2
                  ss1.Text = AdoRs.Fields(0) & ""
                  
                  ss1.Col = 22
                  sREQ_DATE = ss1.Text
                  If sREQ_DATE = "" Then
                     ss1.Text = dtp_end_date.RawData
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
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 23)
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
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 23)
    
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
        ss1.Col = 27
        TXT_CUT_PLT.Text = ss1.Text
    End If
    
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 23)
        
    End If
    
'HYS INSERT START
    Dim sREQ_DATE As String
        ss1.Row = Row
        ss1.Col = 1
        
        If ss1.Text = "1" Then
           COUNT1 = COUNT1 + 1
           If COUNT1 > 36 Then
'              Call MsgBox("一张委托单不能超过36个试样！", vbCritical, "系统提示信息")
              ss1.Text = "0"
              Exit Sub
           End If
           
           ss1.Col = 22
           sREQ_DATE = ss1.Text
           If sREQ_DATE = "" Then
              ss1.Text = dtp_end_date.RawData
           End If
        Else
           Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
        End If
'HYS INSERT END
    
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 23)
    
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
    
    If IsEmpty(arrRecords1) Then
        SAMPLE_SEND_PRINT = "Err Data"
        Exit Function
    End If
    
    
    RowCnt = UBound(arrRecords1, 2)
    
    PrtCnt = -1
    LneCnt = 0
    ROW_NUM = 0
    
    ReDim pAry(1 To 30, 1 To 18)

    
    Do

        LneCnt = LneCnt + 1
        PrtCnt = PrtCnt + 1

        pAry(LneCnt, 1) = ROW_NUM + LneCnt                               ' SEQ
        pAry(LneCnt, 2) = arrRecords1(1, PrtCnt) & ""                    ' SMP_NO
        pAry(LneCnt, 3) = arrRecords1(2, PrtCnt) & ""                    ' STDSPEC
        pAry(LneCnt, 4) = arrRecords1(3, PrtCnt) & ""                    ' THK
        pAry(LneCnt, 5) = arrRecords1(4, PrtCnt) & ""                    ' SMP_CNT
        
        pAry(LneCnt, 6) = arrRecords1(5, PrtCnt) & ""                    ' TENCIL_FL
        pAry(LneCnt, 7) = arrRecords1(6, PrtCnt) & ""                    ' Bend_Fl
        pAry(LneCnt, 8) = arrRecords1(7, PrtCnt) & ""                    ' Impact_TEMP
        pAry(LneCnt, 9) = arrRecords1(8, PrtCnt) & ""                    ' Tim_Imact_TEMP
        pAry(LneCnt, 10) = arrRecords1(9, PrtCnt) & ""                   ' Drop_Wgt_TEMP
        
        pAry(LneCnt, 11) = arrRecords1(10, PrtCnt) & ""                  ' MACRO_FL
        pAry(LneCnt, 12) = arrRecords1(11, PrtCnt) & ""                  ' Non_Metal_Fl
        pAry(LneCnt, 13) = arrRecords1(12, PrtCnt) & ""                  ' Hardness_Fl
        pAry(LneCnt, 14) = arrRecords1(13, PrtCnt) & ""                  ' Chem_Fl
        pAry(LneCnt, 15) = arrRecords1(14, PrtCnt) & ""                  ' Ton_Fl
        pAry(LneCnt, 16) = arrRecords1(15, PrtCnt) & ""                  ' Std_Smp_Fl
        pAry(LneCnt, 17) = arrRecords1(16, PrtCnt) & ""                  ' Photo_Fl
        pAry(LneCnt, 18) = arrRecords1(17, PrtCnt) & ""                  ' Text
        
       
        If LneCnt = 30 Then
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
            
            
             sRow = "A7" & ":R" & 6 + LneCnt
            xlSheet.Range(sRow).Value = pAry
            
            xlApp.Range(sRow).Select
            With xlApp.Selection.Borders
                .LineStyle = 1
            End With
    
            sRow = "A" & LneCnt + 7
            xlApp.Range(sRow).Value = "委托人：" & sUsername
            xlApp.Range(sRow).Font.Size = 10
            sRow = "B" & LneCnt + 7
            xlApp.Range(sRow).Value = "委托时间：" & Format(Now, "YYYY-MM-DD HH:MM:SS")
            xlApp.Range(sRow).Font.Size = 10
            sRow = "F" & LneCnt + 7
            xlApp.Range(sRow).Value = "送样人："
            xlApp.Range(sRow).Font.Size = 10
            sRow = "M" & LneCnt + 7
            xlApp.Range(sRow).Value = "送样时间："
            xlApp.Range(sRow).Font.Size = 10
            
'            xlApp.ActiveSheet.Paste
    
    
            Screen.MousePointer = vbDefault
            xlApp.Application.Visible = True
            Set xlSheet = Nothing
            Set xlApp = Nothing
'            xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
'
'            Set xlSheet = Nothing
'            xlApp.ActiveWorkbook.Close False
'            xlApp.Quit

            LneCnt = 0
            ReDim pAry(1 To 30, 1 To 18)
            
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
         
         sRow = "A7" & ":R" & 6 + LneCnt
        xlSheet.Range(sRow).Value = pAry
        
        xlApp.Range(sRow).Select
        With xlApp.Selection.Borders
            .LineStyle = 1
        End With

        sRow = "A" & LneCnt + 7
        xlApp.Range(sRow).Value = "委托人：" & sUsername
        xlApp.Range(sRow).Font.Size = 10
        sRow = "B" & LneCnt + 7
        xlApp.Range(sRow).Value = "委托时间：" & Format(Now, "YYYY-MM-DD HH:MM:SS")
        xlApp.Range(sRow).Font.Size = 10
        sRow = "F" & LneCnt + 7
        xlApp.Range(sRow).Value = "送样人："
        xlApp.Range(sRow).Font.Size = 10
        sRow = "M" & LneCnt + 7
        xlApp.Range(sRow).Value = "送样时间："
        xlApp.Range(sRow).Font.Size = 10
        
'            xlApp.ActiveSheet.Paste
        
        Screen.MousePointer = vbDefault
        xlApp.Application.Visible = True
        Set xlSheet = Nothing
        Set xlApp = Nothing

'        Set xlSheet = Nothing
'        xlApp.ActiveWorkbook.Close False
'        xlApp.Quit

    End If
    
    Set xlApp = Nothing
    
    Exit Function
    
End Function




