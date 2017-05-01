VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AEB1060C 
   Caption         =   "订单分析结果查询_AEB1060C"
   ClientHeight    =   7620
   ClientLeft      =   600
   ClientTop       =   2835
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7620
   ScaleWidth      =   15225
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_ccm_line 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   6345
      MaxLength       =   1
      TabIndex        =   9
      Tag             =   "连浇机号"
      Top             =   90
      Width           =   390
   End
   Begin VB.TextBox txt_plt_name 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   1650
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "工厂"
      Top             =   90
      Width           =   3135
   End
   Begin VB.TextBox txt_plt 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   1185
      MaxLength       =   2
      TabIndex        =   0
      Tag             =   "工厂"
      Top             =   90
      Width           =   465
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8700
      Left            =   60
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   495
      Width           =   15150
      _ExtentX        =   26723
      _ExtentY        =   15346
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "炼钢相关信息"
      TabPicture(0)   =   "AEB1060C.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ss1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "钢板相关信息"
      TabPicture(1)   =   "AEB1060C.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ss2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "钢卷相关信息"
      TabPicture(2)   =   "AEB1060C.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ss3"
      Tab(2).ControlCount=   1
      Begin FPSpread.vaSpread ss3 
         Height          =   8250
         Left            =   -74940
         TabIndex        =   8
         Top             =   390
         Width           =   15015
         _Version        =   393216
         _ExtentX        =   26485
         _ExtentY        =   14552
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
         MaxCols         =   0
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AEB1060C.frx":0054
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   8250
         Left            =   -74940
         TabIndex        =   7
         Top             =   390
         Width           =   15015
         _Version        =   393216
         _ExtentX        =   26485
         _ExtentY        =   14552
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
         MaxCols         =   0
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AEB1060C.frx":028A
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   8250
         Left            =   60
         TabIndex        =   6
         Top             =   390
         Width           =   15015
         _Version        =   393216
         _ExtentX        =   26485
         _ExtentY        =   14552
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
         MaxCols         =   2
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AEB1060C.frx":04CB
      End
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   135
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
      Left            =   12945
      Top             =   90
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      Caption         =   "钢卷合计"
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
      ForeColor       =   255
   End
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   10620
      Top             =   90
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      Caption         =   "钢板合计"
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
      ForeColor       =   255
   End
   Begin CSTextLibCtl.sidbEdit sdb_coil 
      Height          =   315
      Left            =   13995
      TabIndex        =   4
      Top             =   90
      Width           =   1140
      _Version        =   262145
      _ExtentX        =   2011
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   255
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
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
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   ""
      Text            =   " 0"
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
      NumDecDigits    =   0
      NumIntDigits    =   12
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_plate 
      Height          =   315
      Left            =   11670
      TabIndex        =   3
      Top             =   90
      Width           =   1140
      _Version        =   262145
      _ExtentX        =   2011
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   255
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
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
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   ""
      Text            =   " 0"
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
      NumDecDigits    =   0
      NumIntDigits    =   12
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   8325
      Top             =   90
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      Caption         =   "板坯合计"
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
      ForeColor       =   255
   End
   Begin CSTextLibCtl.sidbEdit sdb_slab 
      Height          =   315
      Left            =   9375
      TabIndex        =   2
      Top             =   90
      Width           =   1140
      _Version        =   262145
      _ExtentX        =   2011
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   255
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
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
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   ""
      Text            =   " 0"
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
      NumDecDigits    =   0
      NumIntDigits    =   12
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel8 
      Height          =   315
      Left            =   5040
      Top             =   90
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Caption         =   "连铸机号"
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
End
Attribute VB_Name = "AEB1060C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       DAILY SCHEDULE
'-- Sub_System Name
'-- Program Name
'-- Program ID        AEB1060C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
'-- Date              2003.7.28
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

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Sc3 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim iSumCol As New Collection       'Sum Column

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"
         
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
         Call Gp_Ms_Collection(txt_plt, "p", "n", "m", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_plt_name, " ", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_ccm_line, "p", "n", "m", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_slab, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_plate, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_coil, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
         
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc2.Add Item:=ss2, Key:="Spread"
    Sc3.Add Item:=ss3, Key:="Spread"
    
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    'Duplicate Count
    'iDupCnt = 1
    
    'Sum Column Count
    'iSumCnt = 1
    
    'Sum Column Setting
    'iSumCol.Add Item:=4
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0

End Sub

Public Sub Sp_Setting()

    ss1.ColWidth(0) = 16
    ss2.ColWidth(0) = 16
    ss3.ColWidth(0) = 16

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
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)
    Call Gp_Sp_Setting(sc2.Item("Spread"), False)
    Call Gp_Sp_Setting(Sc3.Item("Spread"), False)
    
    ss1.MaxCols = 0
    
    Call Sp_Setting
   
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    Call Gf_Sp_Cls(Sc3)
    
    If App.Title = "AE" Then
        txt_plt.Text = "B1"
    ElseIf App.Title = "BE" Then
        txt_plt.Text = "B1"
    End If
    
    Call txt_plt_KeyUp(0, 0)
    txt_ccm_line.Text = "1"

    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Set rControl = Nothing
    
    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set sc2 = Nothing
    Set Sc3 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(sc1) Then
        Call Gf_Sp_Cls(sc2)
        Call Gf_Sp_Cls(Sc3)
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        rControl(1).SetFocus
        
        If App.Title = "AE" Then
            txt_plt.Text = "B1"
        ElseIf App.Title = "BE" Then
            txt_plt.Text = "B1"
        End If
        
        Call txt_plt_KeyUp(0, 0)
        txt_ccm_line.Text = "1"
    End If
    
End Sub

Public Sub Form_Exc()
    
    If SSTab1.Tab = 0 Then
        Call Gp_Sp_Excel(Me, ss1, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    ElseIf SSTab1.Tab = 1 Then
        Call Gp_Sp_Excel(Me, ss2, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    Else
        Call Gp_Sp_Excel(Me, ss3, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    End If
    
End Sub

Public Sub Form_Ref()

    Dim sQuery1 As String   'SLAB  Header Display
    Dim sQuery2 As String   'SLAB  Data Display
    Dim sQuery3 As String   'PLATE Header Display
    Dim sQuery4 As String   'SLAB  Data Display
    Dim sQuery5 As String   'COIL  Header Display
    Dim sQuery6 As String   'COIL  Data Display
    
    Dim sMesg As String
    Dim sCheck As Boolean
    
    sCheck = False

    'SLAB   Header Display
    sQuery1 = "SELECT DISTINCT(SLAB_WID) "
    sQuery1 = sQuery1 + " FROM EP_ORD_EDT "
    sQuery1 = sQuery1 + " WHERE SMS_PLT      = '" + txt_plt.Text + "' "
    sQuery1 = sQuery1 + "   AND SMS_CCM_LINE = '" + txt_ccm_line.Text + "' "
    sQuery1 = sQuery1 + " ORDER BY SLAB_WID "
    
    'SLAB  Data Display
    sQuery2 = "             SELECT STLGRD, SLAB_WID, SUM(NVL(DESIGN_CNF_WGT1,0)),SUM(NVL(DESIGN_CNF_WGT2,0)),SUM(NVL(DESIGN_CNF_WGT3,0)), "
    sQuery2 = sQuery2 + "                            SUM(NVL(DESIGN_CNF_WGT1,0))+SUM(NVL(DESIGN_CNF_WGT2,0))+SUM(NVL(DESIGN_CNF_WGT3,0))"
    sQuery2 = sQuery2 + "     FROM ((SELECT STLGRD, SLAB_WID, SUM(NVL(DESIGN_CNF_WGT,0)) DESIGN_CNF_WGT1,0 DESIGN_CNF_WGT2, 0 DESIGN_CNF_WGT3 "
    sQuery2 = sQuery2 + "              FROM EP_ORD_EDT "
    sQuery2 = sQuery2 + "             WHERE SMS_PLT      = '" + txt_plt.Text + "' "
    sQuery2 = sQuery2 + "               AND SMS_CCM_LINE = '" + txt_ccm_line.Text + "' "
    sQuery2 = sQuery2 + "               AND PROD_CD = 'PP' "
    sQuery2 = sQuery2 + "             GROUP BY STLGRD, SLAB_WID) "
    sQuery2 = sQuery2 + "             UNION ALL "
    sQuery2 = sQuery2 + "           (SELECT STLGRD, SLAB_WID, 0 DESIGN_CNF_WGT1, SUM(NVL(DESIGN_CNF_WGT,0)) DESIGN_CNF_WGT2, 0 DESIGN_CNF_WGT3 "
    sQuery2 = sQuery2 + "              FROM EP_ORD_EDT "
    sQuery2 = sQuery2 + "             WHERE SMS_PLT      = '" + txt_plt.Text + "' "
    sQuery2 = sQuery2 + "               AND SMS_CCM_LINE = '" + txt_ccm_line.Text + "' "
    sQuery2 = sQuery2 + "               AND PROD_CD = 'HC' "
    sQuery2 = sQuery2 + "             GROUP BY STLGRD, SLAB_WID) "
    sQuery2 = sQuery2 + "             UNION ALL "
    sQuery2 = sQuery2 + "           (SELECT STLGRD, SLAB_WID, 0 DESIGN_CNF_WGT1, 0 DESIGN_CNF_WGT2, SUM(NVL(DESIGN_CNF_WGT,0)) DESIGN_CNF_WGT3 "
    sQuery2 = sQuery2 + "              FROM EP_ORD_EDT "
    sQuery2 = sQuery2 + "             WHERE SMS_PLT      = '" + txt_plt.Text + "' "
    sQuery2 = sQuery2 + "               AND SMS_CCM_LINE = '" + txt_ccm_line.Text + "' "
    sQuery2 = sQuery2 + "               AND PROD_CD = 'SL' "
    sQuery2 = sQuery2 + "             GROUP BY STLGRD, SLAB_WID)) "
    sQuery2 = sQuery2 + "             GROUP BY STLGRD, SLAB_WID "
    sQuery2 = sQuery2 + "             ORDER BY STLGRD, SLAB_WID "
    
    'PLATE  Header Display
    sQuery3 = "SELECT DISTINCT(PROD_WID) "
    sQuery3 = sQuery3 + "  FROM EP_ORD_EDT "
    sQuery3 = sQuery3 + " WHERE SMS_PLT      = '" + txt_plt.Text + "' "
    sQuery3 = sQuery3 + "   AND SMS_CCM_LINE = '" + txt_ccm_line.Text + "' "
    sQuery3 = sQuery3 + "   AND PROD_CD = 'PP' ORDER BY PROD_WID "
    
    'PLATE  Data Display
    sQuery4 = "SELECT PROD_THK, PROD_WID, SUM(NVL(DESIGN_CNF_WGT,0)) "
    sQuery4 = sQuery4 + " FROM EP_ORD_EDT "
    sQuery4 = sQuery4 + " WHERE SMS_PLT      = '" + txt_plt.Text + "' "
    sQuery4 = sQuery4 + "   AND SMS_CCM_LINE = '" + txt_ccm_line.Text + "' "
    sQuery4 = sQuery4 + "   AND PROD_CD      = 'PP' "
    sQuery4 = sQuery4 + " GROUP BY PROD_THK, PROD_WID "
    sQuery4 = sQuery4 + " ORDER BY PROD_THK, PROD_WID "
    
    'COIL   Header Display
    sQuery5 = "SELECT DISTINCT(PROD_WID) "
    sQuery5 = sQuery5 + "  FROM EP_ORD_EDT  "
    sQuery5 = sQuery5 + " WHERE SMS_PLT      = '" + txt_plt.Text + "' "
    sQuery5 = sQuery5 + "   AND SMS_CCM_LINE = '" + txt_ccm_line.Text + "' "
    sQuery5 = sQuery5 + "   AND PROD_CD = 'HC' ORDER BY PROD_WID "
    
    'COIL   Data Display
    sQuery6 = "SELECT PROD_THK, PROD_WID, SUM(NVL(DESIGN_CNF_WGT,0)) "
    sQuery6 = sQuery6 + " FROM EP_ORD_EDT "
    sQuery6 = sQuery6 + " WHERE SMS_PLT      = '" + txt_plt.Text + "' "
    sQuery6 = sQuery6 + "   AND SMS_CCM_LINE = '" + txt_ccm_line.Text + "' "
    sQuery6 = sQuery6 + "   AND PROD_CD      = 'HC' "
    sQuery6 = sQuery6 + " GROUP BY PROD_THK, PROD_WID "
    sQuery6 = sQuery6 + " ORDER BY PROD_THK, PROD_WID "
    
    sMesg = Gf_Ms_NeceCheck(nControl)
    If sMesg = "OK" Then
    
        sMesg = Gf_Ms_NeceCheck2(mControl)
        If sMesg = "OK" Then

            'Header Display
            Call Sp_Header_Refer1(ss1, sQuery1)      'SLAB   Header Display
            Call Sp_Header_Refer2(ss2, sQuery3)      'PLATE  Header Display
            Call Sp_Header_Refer2(ss3, sQuery5)      'COIL   Header Display
            
            'Data Display
            If Sp_Data_Refer1(ss1, sQuery2) Then     'SLAB  Data Display
                sCheck = True
                ss1.Col = ss1.MaxCols
                ss1.Row = ss1.MaxRows
                sdb_slab.VALUE = ss1.Text
                ss1.OperationMode = OperationModeRow
            End If
            
            If Sp_Data_Refer2(ss2, sQuery4) Then     'PLATE  Data Display
                sCheck = True
                ss2.Col = ss2.MaxCols
                ss2.Row = ss2.MaxRows
                sdb_plate.VALUE = IIf(ss2.Text = "", 0, ss2.Text)
                ss2.OperationMode = OperationModeRow
            End If
            
            If Sp_Data_Refer2(ss3, sQuery6) Then     'COIL   Data Display
                sCheck = True
                ss3.Col = ss3.MaxCols
                ss3.Row = ss3.MaxRows
                sdb_coil.VALUE = ss3.Text
                ss3.OperationMode = OperationModeRow
            End If
                
            If sCheck Then
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                Call Gp_Sp_ReadOnlySet(sc1.Item("Spread"))
                Call Gp_Sp_ReadOnlySet(sc2.Item("Spread"))
                Call Gp_Sp_ReadOnlySet(Sc3.Item("Spread"))
            End If
            
        Else
            sMesg = sMesg + "长度不正确"
            Call Gp_MsgBoxDisplay(sMesg)
        End If
    
    Else
        sMesg = sMesg + "必须输入"
        Call Gp_MsgBoxDisplay(sMesg)
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

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)

    If Col = 0 Or ss1.MaxCols = Col Or ss1.MaxCols - 1 = Col _
               Or ss1.MaxCols - 2 = Col Or ss1.MaxCols - 3 = Col _
               Or Row = 0 Or ss1.MaxRows = Row Then Exit Sub
        
    ss1.Col = Col
    ss1.Row = Row
    If ss1.Text = "" Or ss1.Text = "0" Then Exit Sub
    If Col Mod 4 = 0 Then Exit Sub
    
    Unload AEB1010C
    Load AEB1010C
    
    If Col Mod 4 = 1 Then AEB1010C.txt_prod_cd.Text = "PP"
    If Col Mod 4 = 2 Then AEB1010C.txt_prod_cd.Text = "HC"
    If Col Mod 4 = 3 Then AEB1010C.txt_prod_cd.Text = "SL"
    
    AEB1010C.txt_plt.Text = txt_plt.Text
    AEB1010C.txt_prc_line.Text = txt_ccm_line.Text
    AEB1010C.txt_ccm_line.Text = txt_ccm_line.Text
    
    AEB1010C.txt_prod_cd_name.Text = ""
    
    ss1.Col = 0
    ss1.Row = Row
    AEB1010C.TxT_stdgrd.Text = ss1.Text
    
    ss1.Col = Col
    ss1.Row = 0
    AEB1010C.TXT_SlaB_WIDTH_FROM.VALUE = ss1.VALUE
    AEB1010C.TXT_SLAB_WIDTH_TO.VALUE = ss1.VALUE
    
    AEB1010C.Txt_urgnt_fl.Text = ""
    AEB1010C.Txt_urgnt_fl_name.Text = ""
    AEB1010C.txt_del_to.RawData = ""
    AEB1010C.txt_prod_thk_from = 0
    AEB1010C.txt_prod_thk_to = 0
    AEB1010C.txt_prod_wid_from = 0
    AEB1010C.txt_prod_wid_to = 0
    AEB1010C.txt_prod_len_from = 0
    AEB1010C.txt_prod_len_to = 0
    
    AEB1010C.Active_CForm = "AEB1060C"
    AEB1010C.Show
    AEB1010C.SetFocus
    
End Sub

Private Sub ss2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss2_DblClick(ByVal Col As Long, ByVal Row As Long)

    If Col = 0 Or ss2.MaxCols = Col Or Row = 0 Or ss2.MaxRows = Row Then Exit Sub
        
    ss2.Col = Col
    ss2.Row = Row
        
    If ss2.Text = "" Then Exit Sub
    
    Unload AEB1010C
    Load AEB1010C
    
    AEB1010C.txt_plt.Text = txt_plt.Text
    AEB1010C.txt_prc_line.Text = txt_ccm_line.Text
    AEB1010C.txt_ccm_line.Text = txt_ccm_line.Text
    
    AEB1010C.txt_prod_cd.Text = "PP"
    
    AEB1010C.TxT_stdgrd.Text = ""
    AEB1010C.Txt_urgnt_fl.Text = ""
    AEB1010C.Txt_urgnt_fl_name.Text = ""
    AEB1010C.txt_del_to.RawData = ""
    AEB1010C.txt_prod_len_from = 0
    AEB1010C.txt_prod_len_to = 0
    AEB1010C.TXT_SlaB_WIDTH_FROM.VALUE = 0
    AEB1010C.TXT_SLAB_WIDTH_TO.VALUE = 0
    
    ss2.Col = 0
    ss2.Row = Row
    AEB1010C.txt_prod_thk_from = ss2.VALUE
    AEB1010C.txt_prod_thk_to = ss2.VALUE
    
    ss2.Col = Col
    ss2.Row = 0
    AEB1010C.txt_prod_wid_from.VALUE = ss2.VALUE
    AEB1010C.txt_prod_wid_to.VALUE = ss2.VALUE
    
    AEB1010C.Active_CForm = "AEB1060C"
    AEB1010C.Show
    AEB1010C.SetFocus
    
End Sub

Private Sub ss3_DblClick(ByVal Col As Long, ByVal Row As Long)

    If Col = 0 Or ss3.MaxCols = Col Or Row = 0 Or ss3.MaxRows = Row Then Exit Sub
            
    ss3.Col = Col
    ss3.Row = Row
        
    If ss3.Text = "" Then Exit Sub
    
    Unload AEB1010C
    Load AEB1010C
    
    AEB1010C.txt_plt.Text = txt_plt.Text
    AEB1010C.txt_prc_line.Text = txt_ccm_line.Text
    AEB1010C.txt_ccm_line.Text = txt_ccm_line.Text
    
    AEB1010C.txt_prod_cd.Text = "HC"
    
    AEB1010C.TxT_stdgrd.Text = ""
    AEB1010C.Txt_urgnt_fl.Text = ""
    AEB1010C.Txt_urgnt_fl_name.Text = ""
    AEB1010C.txt_del_to.RawData = ""
    AEB1010C.txt_prod_len_from = 0
    AEB1010C.txt_prod_len_to = 0
    AEB1010C.TXT_SlaB_WIDTH_FROM.VALUE = 0
    AEB1010C.TXT_SLAB_WIDTH_TO.VALUE = 0
    
    ss3.Col = 0
    ss3.Row = Row
    AEB1010C.txt_prod_thk_from = ss3.VALUE
    AEB1010C.txt_prod_thk_to = ss3.VALUE
    
    ss3.Col = Col
    ss3.Row = 0
    AEB1010C.txt_prod_wid_from.VALUE = ss3.VALUE
    AEB1010C.txt_prod_wid_to.VALUE = ss3.VALUE
    
    AEB1010C.Active_CForm = "AEB1060C"
    AEB1010C.Show
    AEB1010C.SetFocus
    
End Sub

Private Sub ss3_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    Call Gp_Sp_Sort(sc1.Item("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)

    Call Gp_Sp_Sort(sc2.Item("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss3_Click(ByVal Col As Long, ByVal Row As Long)

    Call Gp_Sp_Sort(Sc3.Item("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_LostFocus()
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss2_LostFocus()
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss3_LostFocus()
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    If Row > 0 Then
        Set Active_Spread = Me.ss1
        MDIMain.Mnu_Sorting.Enabled = False
        PopupMenu MDIMain.PopUp_Spread
        MDIMain.Mnu_Sorting.Enabled = True
    End If
    
End Sub

Private Sub ss2_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    If Row > 0 Then
        Set Active_Spread = Me.ss2
        MDIMain.Mnu_Sorting.Enabled = False
        PopupMenu MDIMain.PopUp_Spread
        MDIMain.Mnu_Sorting.Enabled = True
    End If
    
End Sub

Private Sub ss3_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    If Row > 0 Then
        Set Active_Spread = Me.ss3
        MDIMain.Mnu_Sorting.Enabled = False
        PopupMenu MDIMain.PopUp_Spread
        MDIMain.Mnu_Sorting.Enabled = True
    End If
    
End Sub

Private Sub txt_plt_DblClick()

    Call txt_plt_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_plt_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0001"
        DD.rControl.Add Item:=txt_plt
        DD.rControl.Add Item:=txt_plt_name

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else

        If Len(Trim(txt_plt.Text)) = txt_plt.MaxLength Then
            txt_plt_name.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_plt.Text), 2)
        Else
            txt_plt_name.Text = ""
        End If
    
    End If
    

End Sub

Public Function Sp_Header_Refer1(sPname As Variant, sQuery As String) As Boolean

On Error GoTo SpreadDisplay1_Error

    Dim iCol As Integer
    Dim iCnt As Integer
    Dim iColCnt As Integer
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    Set AdoRs = New ADODB.Recordset
    
    With sPname

        Sp_Header_Refer1 = True
        
        .ReDraw = False
        .MaxRows = 0:  .MaxCols = 0
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            Sp_Header_Refer1 = False
            '.ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 2) + 1 <> 0 Then
        
            .MaxCols = (UBound(ArrayRecords, 2) + 1) * 4
            For iCol = 0 To .MaxCols - 1 Step 4
            
                For iColCnt = 1 To 4
                    
                    .Row = SpreadHeader + (.ColHeaderRows - 2)
                    .Col = iCol + iColCnt
                    
                    If VarType(ArrayRecords(0, iCnt)) = vbNull Then
                        .Text = ""
                    Else
                        .Text = Trim(ArrayRecords(0, iCnt))
                    End If
                    
                    .ColWidth(iCol + iColCnt) = 4.5
    
                    .Col = iCol + iColCnt: .Col2 = iCol + iColCnt
                    .Row = 1: .Row2 = -1
                    .BlockMode = True
                    .CellType = 13      'SS_CELL_TYPE_NUMBER
                    .TypeNumberDecPlaces = 0
                    .TypeNumberMax = 999999999
                    .TypeNumberMin = 0
                    .TypeNumberShowSep = True
                    .TypeNumberLeadingZero = TypeLeadingZeroYes

                    .TypeHAlign = TypeHAlignRight
                    .BlockMode = False
                    
                    .Row = SpreadHeader + (.ColHeaderRows - 1)
                    .Col = iCol + iColCnt
                    
                    Select Case iColCnt
                        Case 1
                            .Text = "PP"
                        Case 2
                            .Text = "HC"
                        Case 3
                            .Text = "SL"
                        Case 4
                            .Text = "合计"
                    End Select
                    
                Next iColCnt
                
                iCnt = iCnt + 1
                
            Next iCol
            
            For iColCnt = 1 To 4
                
                .MaxCols = .MaxCols + 1
                .Col = .MaxCols
                .Row = SpreadHeader + (.ColHeaderRows - 2)
                .Text = "合计(t)"
                .Row = SpreadHeader + (.ColHeaderRows - 1)
                
                Select Case iColCnt
                    Case 1
                        .Text = "PP"
                    Case 2
                        .Text = "HC"
                    Case 3
                        .Text = "SL"
                    Case 4
                        .Text = "合计"
                End Select
                    
                .ColWidth(.Col) = 6
                    
                .Col = .MaxCols: .Col2 = .MaxCols
                .Row = 1: .Row2 = -1
                .BlockMode = True
                .CellType = 13      'SS_CELL_TYPE_NUMBER
                .TypeNumberDecPlaces = 0
                .TypeNumberMax = 999999999
                .TypeNumberMin = 0
                .TypeNumberShowSep = True
                .TypeNumberLeadingZero = TypeLeadingZeroNo
                .TypeHAlign = TypeHAlignRight
                .BlockMode = False
                
            Next iColCnt
            
        End If
        
        .BlockMode = True
        .Col = .MaxCols:  .Col2 = .MaxCols
        .Row = 1: .Row2 = -1
        .ForeColor = &HFF&  '&H00FF0000&
        .BlockMode = False
        
        For iColCnt = 4 To .MaxCols - 4 Step 4
            .BlockMode = True
            .Col = iColCnt:  .Col2 = iColCnt
            .Row = 1: .Row2 = -1
            .ForeColor = &HFF0000
            .BlockMode = False
        Next iColCnt
        
        .BlockMode = True
        .Row = 0
        .Col = 1
        .Row2 = 0
        .Col2 = -1
        .RowMerge = MergeAlways
        .ColMerge = MergeAlways
        .BlockMode = False
        
        .ReDraw = True
        .Refresh
        
        Screen.MousePointer = vbDefault
        
    End With
        
    Exit Function

SpreadDisplay1_Error:
    
    Set AdoRs = Nothing
    ss1.ReDraw = True
    Sp_Header_Refer1 = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay1_Error : " & Error)
    
End Function

Public Function Sp_Header_Refer2(sPname As Variant, sQuery As String) As Boolean

On Error GoTo SpreadDisplay2_Error

    Dim iCol As Integer
    Dim iCnt As Integer
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    Set AdoRs = New ADODB.Recordset
    
    With sPname

        Sp_Header_Refer2 = True
        .ReDraw = False
        .MaxRows = 0:  .MaxCols = 0
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            Sp_Header_Refer2 = False
            '.ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 2) + 1 <> 0 Then
        
            .MaxCols = UBound(ArrayRecords, 2) + 1
            .Row = 0
            For iCol = 0 To .MaxCols - 1
            
                .Col = iCol + 1
                If VarType(ArrayRecords(0, iCnt)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(0, iCnt))
                End If
                
                .ColWidth(iCol + 1) = 6

                .Col = iCol + 1: .Col2 = iCol + 1
                .Row = 1: .Row2 = -1
                .BlockMode = True
                .CellType = 13      'SS_CELL_TYPE_NUMBER
                .TypeNumberDecPlaces = 0
                .TypeNumberMax = 999999999
                .TypeNumberMin = 0
                .TypeNumberShowSep = True
                .TypeNumberLeadingZero = TypeLeadingZeroNo
                .TypeHAlign = TypeHAlignRight
                .BlockMode = False
                
                iCnt = iCnt + 1
                
            Next iCol
                
            .MaxCols = .MaxCols + 1
            .Col = .MaxCols
            .Text = "合计(t)"
            
            .ColWidth(.Col) = 8
                
            .Col = .MaxCols: .Col2 = .MaxCols
            .Row = 1: .Row2 = -1
            .BlockMode = True
            .CellType = 13      'SS_CELL_TYPE_NUMBER
            .TypeNumberDecPlaces = 0
            .TypeNumberMax = 999999999
            .TypeNumberMin = 0
            .TypeNumberShowSep = True
            .TypeNumberLeadingZero = TypeLeadingZeroNo
            .TypeHAlign = TypeHAlignRight
            .BlockMode = False
            
        End If
        
        .BlockMode = True
        .Col = .MaxCols:  .Col2 = .MaxCols
        .Row = 1: .Row2 = -1
        .ForeColor = &HFF&
        .BlockMode = False
        
        .ReDraw = True
        .Refresh
        Screen.MousePointer = vbDefault
        
    End With
    
    Exit Function

SpreadDisplay2_Error:
    
    Set AdoRs = Nothing
    ss1.ReDraw = True
    Sp_Header_Refer2 = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay2_Error : " & Error)
    
End Function

Public Function Sp_Data_Refer1(sPname As Variant, sQuery As String) As Boolean

On Error GoTo SpreadDisplay3_Error

    Dim iCol As Integer
    Dim iRow As Integer
    Dim iCnt As Integer
    
    Dim iBas As Integer
    Dim iCot As Integer
    
    Dim sCol_a As String
    Dim sCol_b As String
    Dim sRText As String
    
    Dim ColSum As Double
    
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    Set AdoRs = New ADODB.Recordset
    
    With sPname

        Sp_Data_Refer1 = True
        .ReDraw = False
        .MaxRows = 0
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            Sp_Data_Refer1 = False
            .ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 2) + 1 <> 0 Then
        
            For iCnt = 0 To UBound(ArrayRecords, 2)

                If iCnt = 0 Or sRText <> Trim(ArrayRecords(0, iCnt)) Then
                    sRText = ArrayRecords(0, iCnt)
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                    .Col = 0
                    .Text = sRText
                End If

                .Row = SpreadHeader + (.ColHeaderRows - 2)
                
                For iCol = 1 To .MaxCols Step 4
                
                    .Col = iCol
                    
                    If .Text = Trim(ArrayRecords(1, iCnt)) Then

                        .Row = .MaxRows
                        
                        If VarType(ArrayRecords(2, iCnt)) = vbNull Then
                            .Text = ""
                        Else
                            .Text = Trim(ArrayRecords(2, iCnt))
                        End If
                        
                        .Col = iCol + 1
                        If VarType(ArrayRecords(3, iCnt)) = vbNull Then
                            .Text = ""
                        Else
                            .Text = Trim(ArrayRecords(3, iCnt))
                        End If
                        
                        .Col = iCol + 2
                        If VarType(ArrayRecords(4, iCnt)) = vbNull Then
                            .Text = ""
                        Else
                            .Text = Trim(ArrayRecords(4, iCnt))
                        End If
                        
                        .Col = iCol + 3
                        If VarType(ArrayRecords(5, iCnt)) = vbNull Then
                            .Text = ""
                        Else
                            .Text = Trim(ArrayRecords(5, iCnt))
                        End If
                        
                    End If

                Next iCol
                
            Next iCnt
            
        End If
        
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Col = 0
        .Text = "合计(t)"
        
        Call Gp_Sp_EvenRowBackcolor(sPname, 1)
        
        .BlockMode = True
        .Row = .MaxRows:  .Row2 = .MaxRows
        .Col = 1: .Col2 = -1
        .ForeColor = &HFF&
        .BlockMode = False
        
        For iCol = 4 To .MaxCols - 4 Step 4
            .BlockMode = True
            .Col = iCol:  .Col2 = iCol
            .Row = .MaxRows: .Row2 = .MaxRows
            .ForeColor = &HFF0000
            .BlockMode = False
        Next iCol
        
        'Column Sum
        For iCol = 1 To .MaxCols
        
            .Col = iCol
            
            If .Col <= 26 Then
                sCol_a = Chr(.Col + 64)
                .Formula = "sum(" + sCol_a + "1:" + sCol_a & .MaxRows - 1 & ")"
            Else
                iCot = Int(((.Col - 1) / 26))
                iBas = 26 * iCot
                sCol_a = Chr((.Col - iBas) + 64)
                sCol_b = Chr(iCot + 64)
                .Formula = "sum(" + sCol_b + sCol_a + "1:" + sCol_b + sCol_a & .MaxRows - 1 & ")"
            End If
            
        Next iCol
            
        'Row Sum
        For iRow = 1 To .MaxRows - 1
        
            .Row = iRow
            
            ColSum = 0
            For iCol = 1 To .MaxCols - 4 Step 4
            
                .Col = iCol
                If .Text <> "" Then
                    ColSum = ColSum + .VALUE
                End If
                
            Next iCol
            
            .Col = .MaxCols - 3
            .VALUE = ColSum
                        
            ColSum = 0
            For iCol = 2 To .MaxCols - 4 Step 4
            
                .Col = iCol
                If .Text <> "" Then
                    ColSum = ColSum + .VALUE
                End If
                
            Next iCol
            
            .Col = .MaxCols - 2
            .VALUE = ColSum
            
            ColSum = 0
            For iCol = 3 To .MaxCols - 4 Step 4
            
                .Col = iCol
                If .Text <> "" Then
                    ColSum = ColSum + .VALUE
                End If
                
            Next iCol
            
            .Col = .MaxCols - 1
            .VALUE = ColSum
            
            ColSum = 0
            For iCol = 4 To .MaxCols - 4 Step 4
            
                .Col = iCol
                If .Text <> "" Then
                    ColSum = ColSum + .VALUE
                End If
                
            Next iCol
            
            .Col = .MaxCols
            .VALUE = ColSum
            
        Next iRow
        
        .ReDraw = True
        Call Gp_Ms_ControlLock(Mc1("lControl"), True)
        Screen.MousePointer = vbDefault
        
    End With
    
    Exit Function

SpreadDisplay3_Error:
    
    Set AdoRs = Nothing
    Sp_Data_Refer1 = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay3_Error : " & Error)
    
End Function

Public Function Sp_Data_Refer2(sPname As Variant, sQuery As String) As Boolean

On Error GoTo SpreadDisplay4_Error

    Dim iCol As Integer
    Dim iRow As Integer
    Dim iCnt As Integer
    
    Dim iBas As Integer
    Dim iCot As Integer
    
    Dim sCol_a As String
    Dim sCol_b As String
    
    Dim sRText As String
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    Set AdoRs = New ADODB.Recordset
    
    With sPname

        Sp_Data_Refer2 = True
        .ReDraw = False
        .MaxRows = 0
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            Sp_Data_Refer2 = False
            .ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 2) + 1 <> 0 Then
        
            For iCnt = 0 To UBound(ArrayRecords, 2)

                If iCnt = 0 Or sRText <> Trim(ArrayRecords(0, iCnt)) Then
                    sRText = ArrayRecords(0, iCnt)
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                    .Col = 0
                    .Text = sRText
                End If

                .Row = 0
                For iCol = 1 To .MaxCols
                
                    .Col = iCol
                    If .Text = Trim(ArrayRecords(1, iCnt)) Then

                        .Row = .MaxRows
                        If VarType(ArrayRecords(2, iCnt)) = vbNull Then
                            .Text = ""
                        Else
                            .Text = Trim(ArrayRecords(2, iCnt))
                        End If
                        
                    End If

                Next iCol
                
            Next iCnt
            
        End If
        
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Col = 0
        .Text = "合计(t)"
        
        Call Gp_Sp_EvenRowBackcolor(sPname, 1)
        
        .BlockMode = True
        .Row = .MaxRows:  .Row2 = .MaxRows
        .Col = 1: .Col2 = -1
        .ForeColor = &HFF&
        .BlockMode = False
        
        'Column Sum
        For iCol = 1 To .MaxCols
        
            .Col = iCol
            If .Col <= 26 Then
                sCol_a = Chr(.Col + 64)
                .Formula = "sum(" + sCol_a + "1:" + sCol_a & .MaxRows - 1 & ")"
            Else
                iCot = Int(((.Col - 1) / 26))
                iBas = 26 * iCot
                sCol_a = Chr((.Col - iBas) + 64)
                sCol_b = Chr(iCot + 64)
                .Formula = "sum(" + sCol_b + sCol_a + "1:" + sCol_b + sCol_a & .MaxRows - 1 & ")"
            End If
            
        Next iCol
            
        'Row Sum
        .Col = .MaxCols
        
        For iRow = 1 To .MaxRows - 1
        
            .Row = iRow
            sCol_a = Chr(.MaxCols - 1 + 64)
            .Formula = "sum(A" & iRow & ":" & sCol_a & iRow & ")"
            
        Next iRow
        
        .ReDraw = True
        Call Gp_Ms_ControlLock(Mc1("lControl"), True)
        Screen.MousePointer = vbDefault
        
    End With
    
    Exit Function

SpreadDisplay4_Error:
    
    Set AdoRs = Nothing
    Sp_Data_Refer2 = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay4_Error : " & Error)
    
End Function
