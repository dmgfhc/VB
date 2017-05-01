VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AFK2030C 
   Caption         =   "成分判定实绩查询界面_AFK2030C"
   ClientHeight    =   9225
   ClientLeft      =   135
   ClientTop       =   1935
   ClientWidth     =   15225
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15225
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_CARD_STS 
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
      Height          =   315
      Left            =   6735
      TabIndex        =   11
      Top             =   600
      Width           =   495
   End
   Begin VB.ComboBox cbo_prc_line 
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
      ItemData        =   "AFK2030C.frx":0000
      Left            =   4590
      List            =   "AFK2030C.frx":0002
      TabIndex        =   10
      Tag             =   "连铸机号"
      Top             =   255
      Width           =   700
   End
   Begin VB.CommandButton cmd_bot 
      Caption         =   "▼"
      Height          =   255
      Left            =   2790
      TabIndex        =   9
      Top             =   500
      Width           =   375
   End
   Begin VB.CommandButton cmd_top 
      Caption         =   "▲"
      Height          =   255
      Left            =   2790
      TabIndex        =   8
      Top             =   175
      Width           =   375
   End
   Begin VB.TextBox txt_act_stlgrd 
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
      Height          =   315
      Left            =   10845
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   615
      Width           =   1305
   End
   Begin VB.TextBox txt_act_stlgrd_s 
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
      Left            =   12150
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   615
      Width           =   2955
   End
   Begin VB.TextBox txt_dir_stlgrd_s 
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
      Left            =   12150
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   255
      Width           =   2955
   End
   Begin VB.TextBox txt_dir_stlgrd 
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
      Height          =   315
      Left            =   10845
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   255
      Width           =   1305
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8175
      Left            =   60
      TabIndex        =   3
      Top             =   990
      Width           =   15105
      _Version        =   393216
      _ExtentX        =   26644
      _ExtentY        =   14420
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
      MaxRows         =   13
      SpreadDesigner  =   "AFK2030C.frx":0004
   End
   Begin VB.ComboBox cbo_HEAT_OLC_NO 
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
      Left            =   1365
      TabIndex        =   1
      Tag             =   "炉号"
      Top             =   255
      Width           =   1350
   End
   Begin VB.TextBox txt_ELEMENT_DEC 
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
      Height          =   315
      Left            =   6735
      TabIndex        =   0
      Top             =   255
      Width           =   495
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   180
      Top             =   255
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      Caption         =   "炉号"
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
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   5550
      Top             =   255
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      Caption         =   "判定结果"
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
      Left            =   9690
      Top             =   255
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      Caption         =   "目标钢种号"
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
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   9690
      Top             =   615
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      Caption         =   "实际钢种号"
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
      Left            =   3420
      Top             =   255
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      Caption         =   "机号"
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
      Left            =   5550
      Top             =   600
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      Caption         =   "是否确认"
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "（Y：确认   N：未确认）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7260
      TabIndex        =   12
      Top             =   690
      Width           =   2310
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "（Y：合格   N：不合格）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7260
      TabIndex        =   2
      Top             =   345
      Width           =   2310
   End
End
Attribute VB_Name = "AFK2030C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       NISCO Production Management System
'-- Sub_System Name   Steel Making System
'-- Program Name      CHEMISTRY INQUIRY
'-- Program ID        AFK2030C
'-- Designer          GUOLI
'-- Coder             GUOLI
'-- Date              2003.8.18
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-- 1.01  04.02.12 KIM SUNG HO
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
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"
         
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
    Call Gp_Ms_Collection(cbo_HEAT_OLC_NO, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_ELEMENT_DEC, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_CARD_STS, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

    
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
    sc1.Add Item:="AFK2030C.P_REFER", Key:="P-R"
    
    Proc_Sc.Add Item:=sc1, Key:="Sc1"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0

End Sub

Private Sub cmd_bot_Click()
    Dim V_HEAT_NO As String
    
    If Trim(cbo_HEAT_OLC_NO.Text) = "" Then
       Exit Sub
    End If
    
    V_HEAT_NO = Format(cbo_HEAT_OLC_NO - 1, "00000000")
    'Call Form_Cls
    ss1.ClearRange 3, 1, ss1.MaxCols, ss1.MaxRows, True
    Call Gp_Sp_BlockColor(Proc_Sc("Sc1")("Spread"), 3, ss1.MaxCols, 1, ss1.MaxRows)
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    txt_dir_stlgrd.Text = ""
    txt_act_stlgrd.Text = ""
    txt_dir_stlgrd_s.Text = ""
    txt_act_stlgrd_s.Text = ""
    
    cbo_prc_line.Text = Mid(V_HEAT_NO, 3, 1)
    cbo_HEAT_OLC_NO = V_HEAT_NO
    Call Form_Ref
End Sub

Private Sub cmd_top_Click()
    Dim V_HEAT_NO As String
    
    If Trim(cbo_HEAT_OLC_NO.Text) = "" Then
       Exit Sub
    End If
    
    V_HEAT_NO = Format(cbo_HEAT_OLC_NO + 1, "00000000")
    
    'Call Form_Cls
    ss1.ClearRange 3, 1, ss1.MaxCols, ss1.MaxRows, True
    Call Gp_Sp_BlockColor(Proc_Sc("Sc1")("Spread"), 3, ss1.MaxCols, 1, ss1.MaxRows)
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    txt_dir_stlgrd.Text = ""
    txt_act_stlgrd.Text = ""
    txt_dir_stlgrd_s.Text = ""
    txt_act_stlgrd_s.Text = ""
    
    cbo_prc_line.Text = Mid(V_HEAT_NO, 3, 1)
    cbo_HEAT_OLC_NO = V_HEAT_NO
    Call Form_Ref
End Sub

Private Sub cbo_prc_line_Change()

    ss1.ClearRange 3, 1, ss1.MaxCols, ss1.MaxRows, True
    Call Gp_Sp_BlockColor(Proc_Sc("Sc1")("Spread"), 3, ss1.MaxCols, 1, ss1.MaxRows)
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    txt_dir_stlgrd.Text = ""
    txt_act_stlgrd.Text = ""
    txt_dir_stlgrd_s.Text = ""
    txt_act_stlgrd_s.Text = ""
    
    Call Gf_HeatNo_ComboAdd(M_CN1, cbo_HEAT_OLC_NO, "FP_CONRSLT", "SHIFT", Trim(cbo_prc_line.Text))
    If cbo_HEAT_OLC_NO.ListCount <> 0 And Trim(cbo_HEAT_OLC_NO.Text) = "" Then
       cbo_HEAT_OLC_NO.ListIndex = 0
    End If
    
    If cbo_HEAT_OLC_NO.Text <> "" Then Call Form_Ref
    
End Sub

Private Sub cbo_prc_line_Click()

    Call cbo_prc_line_Change
    
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
    
    cbo_prc_line.AddItem "1"
    cbo_prc_line.AddItem "2"
    cbo_prc_line.AddItem "3"
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    Call Form_Define
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_ControlLock(Mc1("lControl"), True)
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Sp_Header_display(Proc_Sc("Sc1")("Spread"))
    Call Sp_Setting(Proc_Sc("Sc1")("Spread"))
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc1")("Spread"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc1")("Spread"), "F-System.INI", Me.Name)
   
    cbo_prc_line.Text = "1"
    
    cbo_HEAT_OLC_NO.Text = Mid(cbo_HEAT_OLC_NO.Text, 1, 8)
    Screen.MousePointer = vbDefault
    'Call Form_Ref
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("Sc1")("Spread"), "F-System.INI", Me.Name)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()

    ss1.ClearRange 3, 1, ss1.MaxCols, ss1.MaxRows, True
    Call Gp_Sp_BlockColor(Proc_Sc("Sc1")("Spread"), 3, ss1.MaxCols, 1, ss1.MaxRows)
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        
    cbo_HEAT_OLC_NO.Enabled = True
    
    cbo_prc_line.Text = "1"
    
    txt_dir_stlgrd.Text = ""
    txt_act_stlgrd.Text = ""
    txt_dir_stlgrd_s.Text = ""
    txt_act_stlgrd_s.Text = ""
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc1")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
  
End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    Dim sMsg, sSampleCode As String
    Dim sQuery As String
    Dim sQuery_cnt As String
    Dim sQuery_A As String
    cbo_HEAT_OLC_NO.Text = Mid(cbo_HEAT_OLC_NO.Text, 1, 8)
    
    sMsg = Gf_Ms_NeceCheck(Mc1("nControl"))
    If sMsg <> "OK" Then
        sMsg = sMsg + "必须输入"
        Call Gp_MsgBoxDisplay(sMsg)
        Exit Sub
    End If
            
    sSampleCode = Gf_FloatFind(M_CN1, "SELECT LAST_SMP_CD FROM FP_CHARGE WHERE HEAT_NO = '" + cbo_HEAT_OLC_NO.Text + "'")
    sQuery = " SELECT COUNT(*) FROM FP_CHEMISTRY WHERE ELEMENT_DEC = 'N' AND HEAT_NO = '" + cbo_HEAT_OLC_NO.Text + "' AND SAMPLE_CD = '" + sSampleCode + "' "
    sQuery_cnt = " SELECT COUNT(*) From QP_CHEM_SEQ ORDER BY CHEM_COMP_SEQ ASC "
    sQuery_A = " SELECT CARD_STS  FROM   QP_CHEM_RSLT   WHERE HEAT_NO = '" + cbo_HEAT_OLC_NO.Text + "' AND CHEM_COMP_CD = 'C'"
    
'
    
    If Sp_Display(M_CN1, Proc_Sc("Sc1")("Spread"), Gf_Ms_MakeQuery(Proc_Sc("Sc1").Item("P-R"), "R", Mc1("pControl"))) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        
        If Gf_FloatFind(M_CN1, sQuery) = 0 Then
            txt_ELEMENT_DEC.Text = "Y"
        Else
            txt_ELEMENT_DEC.Text = "N"
        End If

        If Gf_FloatFind(M_CN1, sQuery_A) = 0 Then
            txt_CARD_STS.Text = "N"
        Else
            txt_CARD_STS.Text = "Y"
        End If
''
        
'        Call Gp_Ms_ControlLock(Mc1("pControl"), True)

        sQuery = "select stlgrd from ep_charge_ins  where heat_mana_no = '" + cbo_HEAT_OLC_NO.Text + "'"
        txt_dir_stlgrd = Gf_FloatFind(M_CN1, sQuery)
        sQuery = "select b.STEEL_GRD_DETAIL from ep_charge_ins a, qp_nisco_chmc b where heat_mana_no = '" + cbo_HEAT_OLC_NO.Text + "' and a.stlgrd = b.stlgrd  "
        txt_dir_stlgrd_s = Gf_FloatFind(M_CN1, sQuery)
        sQuery = "select ACT_STLGRD from fp_charge where heat_no = '" + cbo_HEAT_OLC_NO.Text + "'"
        txt_act_stlgrd = Gf_FloatFind(M_CN1, sQuery)
        sQuery = "select b.STEEL_GRD_DETAIL from fp_charge a, qp_nisco_chmc b where heat_no = '" + cbo_HEAT_OLC_NO.Text + "' and a.ACT_STLGRD = b.stlgrd  "
        txt_act_stlgrd_s = Gf_FloatFind(M_CN1, sQuery)
        
    End If
    
    Call Gp_Color_Display(ss1, Gf_FloatFind(M_CN1, sQuery_cnt))
            
Refer_Err:

End Sub

Public Sub Gp_Color_Display(sPname As Variant, Col_cnt As Variant)

On Error GoTo SpreadDisplay_Error

    Dim iRowCount As Long
    Dim iColcount As Long
    Dim iText As String
    
    Dim Cnt As Integer
    Cnt = Col_cnt
        
    With sPname
       .ReDraw = False
       For iColcount = 3 To Cnt + 2
           .Col = iColcount
               .Row = 5
               iText = Trim(.Text)
               If iText = "Y" Then
                  .Row = 1
                  .ForeColor = &H80000008
                  .Row = 2
                  .ForeColor = &H80000008
               ElseIf iText = "N" Then
                  .Row = 1
                  .ForeColor = &HFF&
                  .Row = 2
                  .ForeColor = &HFF&
                  .Row = 4
                  .ForeColor = &HFF&
               End If
               
               .Row = 7
               iText = Trim(.Text)
                If iText = "Y" Then
                  .Row = 1
                  .ForeColor = &H80000008
                  .Row = 2
                  .ForeColor = &H80000008
                ElseIf iText = "N" Then
                  .Row = 1
                  .ForeColor = &HFF&
                  .Row = 2
                  .ForeColor = &HFF&
                  .Row = 6
                  .ForeColor = &HFF&
                End If
        
               .Row = 9
               iText = Trim(.Text)
                If iText = "Y" Then
                  .Row = 1
                  .ForeColor = &H80000008
                  .Row = 2
                  .ForeColor = &H80000008
                ElseIf iText = "N" Then
                  .Row = 1
                  .ForeColor = &HFF&
                  .Row = 2
                  .ForeColor = &HFF&
                  .Row = 8
                  .ForeColor = &HFF&
                End If
                
               .Row = 11
               iText = Trim(.Text)
                If iText = "Y" Then
                  .Row = 1
                  .ForeColor = &H80000008
                  .Row = 2
                  .ForeColor = &H80000008
                ElseIf iText = "N" Then
                  .Row = 1
                  .ForeColor = &HFF&
                  .Row = 2
                  .ForeColor = &HFF&
                  .Row = 10
                  .ForeColor = &HFF&
                End If
                
               .Row = 13
               iText = Trim(.Text)
                If iText = "Y" Then
                  .Row = 1
                  .ForeColor = &H80000008
                  .Row = 2
                  .ForeColor = &H80000008
                ElseIf iText = "N" Then
                  .Row = 1
                  .ForeColor = &HFF&
                  .Row = 2
                  .ForeColor = &HFF&
                  .Row = 12
                  .ForeColor = &HFF&
                End If
       Next iColcount
       .ReDraw = True
 
    End With
    
SpreadDisplay_Error:

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

Private Sub ss1_LostFocus()
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    MDIMain.Mnu_Sorting.Enabled = False

    If Row > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If

    MDIMain.Mnu_Sorting.Enabled = True
    
End Sub

Public Sub Sp_Setting(ByVal sPname As Variant)

    Dim iRow As Integer

    With sPname
    
        .RowHeight(-1) = 14
        
        If .ColHeaderRows > 1 Then
            .RowHeight(SpreadHeader + (.ColHeaderRows - 2)) = 12
            .RowHeight(SpreadHeader + (.ColHeaderRows - 1)) = 12
        Else
            .RowHeight(0) = 24
        End If
        
        .RowHeadersShow = False
        .BackColorStyle = BackColorStyleUnderGrid
        
        .GrayAreaBackColor = &HE0E0E0
        .GridColor = &H808040
        
        .ShadowColor = &HE1E4CD
        .ShadowDark = &H808040
        .SelBackColor = &HCEECFF     ''&HE3F4FF      ''&HFFFF80     '&H808040
     
        .OperationMode = OperationModeNormal
        .RetainSelBlock = True
        .UserResize = UserResizeColumns
        
        .ProcessTab = True
        .ScrollBarExtMode = True
        .ButtonDrawMode = 1
        .TabStop = False
        
        .Col = 0: .Col2 = -1
        .Row = 0: .Row2 = -1
        
        .BlockMode = True
        .FontBold = False
        .FontName = "SimSun"
        .FontSize = 10
        .BlockMode = False
        
        .Col = -1
        .Row = 0
        .FontBold = True
        
        For iRow = 1 To .MaxRows
            
            .Col = 3: .Col2 = .MaxCols
            .Row = iRow: .Row2 = iRow
            .BlockMode = True
                    
            Select Case iRow
                Case 1, 2, 3, 4, 6, 8, 10, 12
                    .CellType = CellTypeNumber
                    .TypeNumberDecPlaces = 4
                    .TypeNumberMax = 99.9999
                    .TypeNumberMin = 0
                    .TypeNumberLeadingZero = TypeLeadingZeroYes
                    .TypeHAlign = TypeHAlignRight
                    .TypeVAlign = TypeVAlignCenter
                Case Else
                    .CellType = CellTypeEdit
                    .TypeHAlign = SS_CELL_H_ALIGN_CENTER
                    .TypeVAlign = TypeVAlignCenter
            End Select
            
            .BlockMode = False
                    
        Next iRow
        
    End With
    
End Sub

Public Sub Sp_Header_display(sPname As Variant)

On Error GoTo SpreadDisplay_Error

    Dim iCol As Integer
    Dim iCnt As Integer
    Dim iColCnt As Integer
    Dim sQuery As String
    
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    Set AdoRs = New ADODB.Recordset
    
    sQuery = " SELECT CHEM_COMP_CD From QP_CHEM_SEQ ORDER BY CHEM_COMP_SEQ ASC "
    
    With sPname

        .ReDraw = False
        .MaxCols = 2
        Screen.MousePointer = vbHourglass
        
        'Title Setting
        .Col = 1
        .Row = 0
        .Text = "工序\成分"
        
        .Row = 1
        .Text = "标准"
        .Row = 2
        .Text = "标准"
        .Row = 3
        .Text = "标准"
        .Row = 4
        .Text = "转炉"
        .Row = 5
        .Text = "转炉"
        .Row = 6
        .Text = "LF"
        .Row = 7
        .Text = "LF"
        .Row = 8
        .Text = "VD"
        .Row = 9
        .Text = "VD"
        .Row = 10
        .Text = "RH"
        .Row = 11
        .Text = "RH"
        .Row = 12
        .Text = "中间罐"
        .Row = 13
        .Text = "中间罐"
        
        .Col = 2
        .Row = 0
        .Text = "工序\成分"
        .Row = 1
        .Text = "最小值"
        .Row = 2
        .Text = "最大值"
        .Row = 3
        .Text = "目标值"
        .Row = 4
        .Text = "实绩"
        .Row = 5
        .Text = "判定"
        .Row = 6
        .Text = "实绩"
        .Row = 7
        .Text = "判定"
        .Row = 8
        .Text = "实绩"
        .Row = 9
        .Text = "判定"
        .Row = 10
        .Text = "实绩"
        .Row = 11
        .Text = "判定"
        .Row = 12
        .Text = "实绩"
        .Row = 13
        .Text = "判定"
        
        'Ado Execute
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
            AdoRs.Close
            Set AdoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
        ArrayRecords = AdoRs.GetRows
        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 2) + 1 <> 0 Then
        
            .MaxCols = (UBound(ArrayRecords, 2) + 1) + 2
            .Row = 0
        
            For iCol = 2 To .MaxCols - 1
            
                .Col = iCol + 1
                .ColWidth(.Col) = 8
                
                If VarType(ArrayRecords(0, iCol - 2)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(0, iCol - 2))
                End If
                    
            Next iCol
            
        End If
        
        .BlockMode = True
        .Row = 0
        .Col = 1
        .Row2 = -1
        .Col2 = 2
        .RowMerge = MergeAlways
        .ColMerge = MergeAlways
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
        .BlockMode = False
        
        .ColsFrozen = 2
        .ReDraw = True
        .Refresh
        
        Screen.MousePointer = vbDefault
        
    End With
    
Exit Sub

SpreadDisplay_Error:
    
    Set AdoRs = Nothing
    ss1.ReDraw = True
    Screen.MousePointer = vbDefault
    
End Sub

Public Function Sp_Display(Conn As ADODB.Connection, sPname As Variant, sQuery As String) As Boolean

On Error GoTo SpreadDisplay_Error

    Dim iCount As Integer
    Dim iRowCount As Long
    Dim iColcount As Long
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    Sp_Display = True
    
    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then Sp_Display = False: Exit Function
    End If
    
    Set AdoRs = New ADODB.Recordset
    
    With sPname

        .ReDraw = False
        iCount = 0
        
        .ClearRange 3, 1, .MaxCols, .MaxRows, True
        Call Gp_Sp_BlockColor(Proc_Sc("Sc1")("Spread"), 3, .MaxCols, 1, .MaxRows)
    
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, Conn, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
            
            .ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
            Sp_Display = False
            Call Gp_MsgBoxDisplay("无相关记录", "I")
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        
        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 2) <> 0 Then
        
            For iColcount = 2 To .MaxCols - 1
            
                .Col = iColcount + 1
                
                For iRowCount = 1 To .MaxRows
                
                    .Row = iRowCount
                    
                    If VarType(ArrayRecords(iRowCount, iColcount - 2)) = vbNull Then
                        .Text = ""
                    Else
                        .Text = Trim(ArrayRecords(iRowCount, iColcount - 2))
                    End If
                    
                Next iRowCount
                
            Next iColcount
            
        End If
        
        .ReDraw = True
        Screen.MousePointer = vbDefault
        
    End With

Exit Function

SpreadDisplay_Error:
    
    Set AdoRs = Nothing
    Sp_Display = False
    Call Gp_MsgBoxDisplay("Query Failed..." & sQuery)
    Screen.MousePointer = vbDefault

End Function

Private Sub txt_dir_stlgrd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sQuery As String
    If Trim(txt_dir_stlgrd) <> "" Then
        sQuery = "select steel_grd_detail from qp_stlgrd_inf where stlgrd = '" + txt_dir_stlgrd + "'"
        txt_dir_stlgrd.ToolTipText = Gf_FloatFind(M_CN1, sQuery)
    Else
        txt_dir_stlgrd.ToolTipText = ""
    End If
End Sub
