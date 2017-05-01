VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form ACA1140C 
   Caption         =   "物料工艺时间查询_ACA1140C"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_mode_fl 
      Height          =   285
      Left            =   8280
      MaxLength       =   1
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3360
      TabIndex        =   1
      Top             =   0
      Width           =   4665
      Begin VB.OptionButton opt_dzb 
         BackColor       =   &H00E0E0E0&
         Caption         =   "DZB当前"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   240
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   170
         Value           =   -1  'True
         Width           =   1200
      End
      Begin VB.OptionButton opt_qab 
         BackColor       =   &H00E0E0E0&
         Caption         =   "QAB当前"
         Height          =   195
         Left            =   1740
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   170
         Width           =   1200
      End
      Begin VB.OptionButton opt_his 
         BackColor       =   &H00E0E0E0&
         Caption         =   "历史查询"
         Height          =   195
         Left            =   3360
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   170
         Width           =   1200
      End
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   360
      Top             =   120
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Caption         =   "生产时间"
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
   Begin InDate.UDate txt_mill_occr_date 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Tag             =   "交货期"
      Top             =   120
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Text            =   "____-__"
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
      Mask            =   "%%%%-%%"
      MaxLength       =   7
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8625
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   15915
      _Version        =   393216
      _ExtentX        =   28072
      _ExtentY        =   15214
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
      MaxCols         =   73
      MaxRows         =   1
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "ACA1140C.frx":0000
   End
   Begin Threed.SSPanel SSP90 
      Height          =   375
      Left            =   9360
      TabIndex        =   7
      Top             =   45
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
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
      Caption         =   "该报表最终解释权归生产部所有"
      FloodColor      =   65535
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
End
Attribute VB_Name = "ACA1140C"
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
'-- Program ID        ACA1140C
'-- Document No       Q-00-0010(Specification)
'-- Designer          CaoLei
'-- Coder             CaoLei
'-- Date              2013.10.08
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

Const SS1_ORD_FL = 7                '订单材标记
Const SS1_CUST_CD = 20              '当前客户编码
Const SS1_ORG_ORD_NO = 21           '原始订单号
Const SS1_ORG_ORD_ITEM = 22         '原始订单项次
Const SS1_INSP_OCCR_DATE = 23       '表判合格时间
Const SS1_INSP_OCCR_DATE_1 = 24     '具备入库时间

Const SS1_F_GAS_FL = 25             '切割标记1
Const SS1_F_GAS_RSLT = 26           '切割标记2
Const SS1_GAS_TARD_DATE = 31        '切割拖期时间

Const SS1_F_UST_FL = 32             '探伤标记1
Const SS1_F_UST_RSLT = 33           '探伤标记2
Const SS1_UST_TARD_DATE = 38        '探伤拖期时间

Const SS1_F_CL_FL = 39              '矫直标记1
Const SS1_F_CL_RSLT = 40            '矫直标记2
Const SS1_CL_TARD_DATE = 45         '矫直拖期时间

Const SS1_F_HTM_FL = 47             '热处理标记2
Const SS1_F_HTM_REQ_TARD = 53       '是否热处理条件拖期
Const SS1_HTM_REQ_TARD_DATE = 54    '热处理条件拖期时间
Const SS1_F_HTM_TARD = 55           '是否热处理拖期
Const SS1_HTM_TARD_DATE = 56        '热处理拖期时间
Const SS1_OTHER_TARD = 57           '其他拖期
Const SS1_DZB_DATE = 58             '精整处理时间

Const SS1_CERT_TYPE = 59            '质保书类型
Const SS1_L2_SND_DATE = 64          '委托时间
Const SS1_QAB_W_DATE = 65           '性能处理时间
Const SS1_QAB_DATE = 66             '性能判定时间
Const SS1_RE_TEST_FL = 67           '判定保留
Const SS1_INSP_CD = 68              '检查机关
Const SS1_OTHER_INSP_CD = 69        '检查机关123
Const SS1_COLOR_STROKE = 70         '订单备注

Private Sub Form_Define()

   Dim i As Integer
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
 '   FormType = "Msheet"
    FormType = "Msheet"

   'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
   Call Gp_Ms_Collection(txt_mill_occr_date, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_mode_fl, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          
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
     For i = 1 To ss1.MaxCols
        Call Gp_Sp_Collection(ss1, i, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Next i
     
   'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"

    sc1.Add Item:="ACA1140C.P_SREFER", Key:="P-R"
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
    sc1.Item("Spread").Text = "◎"
    
       
        
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
    
    
    
    txt_mill_occr_date.Text = Mid(Date, 1, 7)    '本月
    
    
    txt_mode_fl.Text = "D"

    
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
    
    
    
    txt_mill_occr_date.Text = Mid(Date, 1, 7)    '本月
   
   
End Sub

Public Sub Form_Exc()

    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    
End Sub

Public Sub Form_Ref()
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        ss1.OperationMode = OperationModeNormal
    End If
    
    If txt_mode_fl.Text = "D" Then
       Call Gp_Sp_ColHidden(ss1, SS1_ORD_FL, False)
       Call Gp_Sp_ColHidden(ss1, SS1_CUST_CD, False)
       Call Gp_Sp_ColHidden(ss1, SS1_ORG_ORD_NO, False)
       Call Gp_Sp_ColHidden(ss1, SS1_ORG_ORD_ITEM, False)
       Call Gp_Sp_ColHidden(ss1, SS1_F_GAS_FL, False)
       Call Gp_Sp_ColHidden(ss1, SS1_F_GAS_RSLT, False)
       Call Gp_Sp_ColHidden(ss1, SS1_GAS_TARD_DATE, False)
       Call Gp_Sp_ColHidden(ss1, SS1_F_UST_FL, False)
       Call Gp_Sp_ColHidden(ss1, SS1_F_UST_RSLT, False)
       Call Gp_Sp_ColHidden(ss1, SS1_UST_TARD_DATE, False)
       Call Gp_Sp_ColHidden(ss1, SS1_F_CL_FL, False)
       Call Gp_Sp_ColHidden(ss1, SS1_F_CL_RSLT, False)
       Call Gp_Sp_ColHidden(ss1, SS1_CL_TARD_DATE, False)
       Call Gp_Sp_ColHidden(ss1, SS1_F_HTM_FL, False)
'       Call Gp_Sp_ColHidden(ss1, SS1_F_HTM_REQ_TARD, False)
       Call Gp_Sp_ColHidden(ss1, SS1_HTM_REQ_TARD_DATE, False)
       Call Gp_Sp_ColHidden(ss1, SS1_F_HTM_TARD, False)
       Call Gp_Sp_ColHidden(ss1, SS1_HTM_TARD_DATE, False)
       Call Gp_Sp_ColHidden(ss1, SS1_OTHER_TARD, False)
       Call Gp_Sp_ColHidden(ss1, SS1_DZB_DATE, False)

       Call Gp_Sp_ColHidden(ss1, SS1_INSP_OCCR_DATE, True)
       Call Gp_Sp_ColHidden(ss1, SS1_INSP_OCCR_DATE_1, True)
'       Call Gp_Sp_ColHidden(ss1, SS1_CERT_TYPE, True)
'       Call Gp_Sp_ColHidden(ss1, SS1_L2_SND_DATE, True)
'       Call Gp_Sp_ColHidden(ss1, SS1_QAB_W_DATE, True)
'       Call Gp_Sp_ColHidden(ss1, SS1_QAB_DATE, True)

'       Call Gp_Sp_ColHidden(ss1, SS1_RE_TEST_FL, True)
'       Call Gp_Sp_ColHidden(ss1, SS1_INSP_CD, True)
'       Call Gp_Sp_ColHidden(ss1, SS1_OTHER_INSP_CD, True)
'       Call Gp_Sp_ColHidden(ss1, SS1_COLOR_STROKE, False)

    ElseIf txt_mode_fl.Text = "Q" Then

       Call Gp_Sp_ColHidden(ss1, SS1_ORD_FL, False)
       Call Gp_Sp_ColHidden(ss1, SS1_INSP_OCCR_DATE, False)
       Call Gp_Sp_ColHidden(ss1, SS1_INSP_OCCR_DATE_1, False)
'       Call Gp_Sp_ColHidden(ss1, SS1_CERT_TYPE, False)
'       Call Gp_Sp_ColHidden(ss1, SS1_L2_SND_DATE, False)
'       Call Gp_Sp_ColHidden(ss1, SS1_QAB_W_DATE, False)
'       Call Gp_Sp_ColHidden(ss1, SS1_QAB_DATE, False)
'       Call Gp_Sp_ColHidden(ss1, SS1_RE_TEST_FL, False)
'       Call Gp_Sp_ColHidden(ss1, SS1_INSP_CD, False)
'       Call Gp_Sp_ColHidden(ss1, SS1_OTHER_INSP_CD, False)
'       Call Gp_Sp_ColHidden(ss1, SS1_COLOR_STROKE, False)

       Call Gp_Sp_ColHidden(ss1, SS1_CUST_CD, True)
       Call Gp_Sp_ColHidden(ss1, SS1_ORG_ORD_NO, True)
       Call Gp_Sp_ColHidden(ss1, SS1_ORG_ORD_ITEM, True)
       Call Gp_Sp_ColHidden(ss1, SS1_F_GAS_FL, True)
       Call Gp_Sp_ColHidden(ss1, SS1_F_GAS_RSLT, True)
       Call Gp_Sp_ColHidden(ss1, SS1_GAS_TARD_DATE, True)
       Call Gp_Sp_ColHidden(ss1, SS1_F_UST_FL, True)
       Call Gp_Sp_ColHidden(ss1, SS1_F_UST_RSLT, True)
       Call Gp_Sp_ColHidden(ss1, SS1_UST_TARD_DATE, True)
       Call Gp_Sp_ColHidden(ss1, SS1_F_CL_FL, True)
       Call Gp_Sp_ColHidden(ss1, SS1_F_CL_RSLT, True)
       Call Gp_Sp_ColHidden(ss1, SS1_CL_TARD_DATE, True)
       Call Gp_Sp_ColHidden(ss1, SS1_F_HTM_FL, True)
'       Call Gp_Sp_ColHidden(ss1, SS1_F_HTM_REQ_TARD, True)
       Call Gp_Sp_ColHidden(ss1, SS1_HTM_REQ_TARD_DATE, True)
       Call Gp_Sp_ColHidden(ss1, SS1_F_HTM_TARD, True)
       Call Gp_Sp_ColHidden(ss1, SS1_HTM_TARD_DATE, True)
       Call Gp_Sp_ColHidden(ss1, SS1_OTHER_TARD, True)
       Call Gp_Sp_ColHidden(ss1, SS1_DZB_DATE, True)


    Else

       Call Gp_Sp_ColHidden(ss1, SS1_INSP_OCCR_DATE, False)
       Call Gp_Sp_ColHidden(ss1, SS1_INSP_OCCR_DATE_1, False)
'       Call Gp_Sp_ColHidden(ss1, SS1_CERT_TYPE, False)
'       Call Gp_Sp_ColHidden(ss1, SS1_L2_SND_DATE, False)
'       Call Gp_Sp_ColHidden(ss1, SS1_QAB_W_DATE, False)
'       Call Gp_Sp_ColHidden(ss1, SS1_QAB_DATE, False)
'       Call Gp_Sp_ColHidden(ss1, SS1_RE_TEST_FL, False)
'       Call Gp_Sp_ColHidden(ss1, SS1_INSP_CD, False)
'       Call Gp_Sp_ColHidden(ss1, SS1_OTHER_INSP_CD, False)
'       Call Gp_Sp_ColHidden(ss1, SS1_COLOR_STROKE, False)

       Call Gp_Sp_ColHidden(ss1, SS1_ORD_FL, True)
       Call Gp_Sp_ColHidden(ss1, SS1_CUST_CD, True)
       Call Gp_Sp_ColHidden(ss1, SS1_ORG_ORD_NO, True)
       Call Gp_Sp_ColHidden(ss1, SS1_ORG_ORD_ITEM, True)
       Call Gp_Sp_ColHidden(ss1, SS1_F_GAS_FL, True)
       Call Gp_Sp_ColHidden(ss1, SS1_F_GAS_RSLT, True)
       Call Gp_Sp_ColHidden(ss1, SS1_GAS_TARD_DATE, True)
       Call Gp_Sp_ColHidden(ss1, SS1_F_UST_FL, True)
       Call Gp_Sp_ColHidden(ss1, SS1_F_UST_RSLT, True)
       Call Gp_Sp_ColHidden(ss1, SS1_UST_TARD_DATE, True)
       Call Gp_Sp_ColHidden(ss1, SS1_F_CL_FL, True)
       Call Gp_Sp_ColHidden(ss1, SS1_F_CL_RSLT, True)
       Call Gp_Sp_ColHidden(ss1, SS1_CL_TARD_DATE, True)
       Call Gp_Sp_ColHidden(ss1, SS1_F_HTM_FL, True)
'       Call Gp_Sp_ColHidden(ss1, SS1_F_HTM_REQ_TARD, True)
       Call Gp_Sp_ColHidden(ss1, SS1_HTM_REQ_TARD_DATE, True)
       Call Gp_Sp_ColHidden(ss1, SS1_F_HTM_TARD, True)
       Call Gp_Sp_ColHidden(ss1, SS1_HTM_TARD_DATE, True)
       Call Gp_Sp_ColHidden(ss1, SS1_OTHER_TARD, True)
       Call Gp_Sp_ColHidden(ss1, SS1_DZB_DATE, True)

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

Private Sub opt_dzb_Click()
    If opt_dzb.Value = True Then
       opt_dzb.ForeColor = &HFF&
       opt_qab.ForeColor = &H80000012
       opt_his.ForeColor = &H80000012
       txt_mode_fl.Text = "D"
    Else
       opt_qab.ForeColor = &HFF&
       opt_dzb.ForeColor = &H80000012
       txt_mode_fl.Text = "Q"
    End If

End Sub

Private Sub opt_his_Click()
    If opt_his.Value = True Then
       opt_his.ForeColor = &HFF&
       opt_dzb.ForeColor = &H80000012
       opt_qab.ForeColor = &H80000012
       txt_mode_fl.Text = "H"
    End If
End Sub

Private Sub opt_qab_Click()
    If opt_qab.Value = True Then
       opt_qab.ForeColor = &HFF&
       opt_dzb.ForeColor = &H80000012
       opt_his.ForeColor = &H80000012
       txt_mode_fl.Text = "Q"
    Else
       opt_dzb.ForeColor = &HFF&
       opt_qab.ForeColor = &H80000012
       txt_mode_fl.Text = "D"
       
    End If
End Sub
