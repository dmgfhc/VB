VERSION 5.00
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form CGF2040C 
   Caption         =   "轧辊/轴承座和轴承的库存管理界面_CGF2040C"
   ClientHeight    =   8265
   ClientLeft      =   180
   ClientTop       =   1725
   ClientWidth     =   14805
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_roll_sts 
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
      TabIndex        =   7
      Tag             =   "供货商"
      Top             =   90
      Width           =   1215
   End
   Begin VB.TextBox TXT_ROLL_MAKER 
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
      Left            =   9585
      TabIndex        =   6
      Tag             =   "供货商"
      Top             =   90
      Width           =   1215
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8745
      Left            =   60
      TabIndex        =   2
      Top             =   495
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   15425
      _Version        =   196609
      PaneTree        =   "CGF2040C.frx":0000
      Begin FPSpread.vaSpread ss1 
         Height          =   8685
         Left            =   30
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   30
         Width           =   8985
         _Version        =   393216
         _ExtentX        =   15849
         _ExtentY        =   15319
         _StockProps     =   64
         ButtonDrawMode  =   4
         ColHeaderDisplay=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   20
         ProcessTab      =   -1  'True
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CGF2040C.frx":0072
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   8685
         Left            =   9105
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   30
         Width           =   3135
         _Version        =   393216
         _ExtentX        =   5530
         _ExtentY        =   15319
         _StockProps     =   64
         ButtonDrawMode  =   4
         ColHeaderDisplay=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   9
         ProcessTab      =   -1  'True
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CGF2040C.frx":2074
      End
      Begin FPSpread.vaSpread ss3 
         Height          =   8685
         Left            =   12330
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   30
         Width           =   2805
         _Version        =   393216
         _ExtentX        =   4948
         _ExtentY        =   15319
         _StockProps     =   64
         ButtonDrawMode  =   4
         ColHeaderDisplay=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   9
         ProcessTab      =   -1  'True
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CGF2040C.frx":3CA3
      End
   End
   Begin VB.ComboBox CBO_ROLL_ID 
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
      Left            =   1230
      TabIndex        =   0
      Top             =   90
      Width           =   1335
   End
   Begin VB.ComboBox CBO_PLT 
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
      ItemData        =   "CGF2040C.frx":58A7
      Left            =   3915
      List            =   "CGF2040C.frx":58AE
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   690
   End
   Begin InDate.ULabel ULabel16 
      Height          =   315
      Left            =   105
      Top             =   90
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "轧辊号"
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
      Left            =   2790
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "工厂代码"
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
   Begin InDate.ULabel ULabel19 
      Height          =   315
      Left            =   8235
      Top             =   90
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      Caption         =   "供货商"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   4140
      Top             =   90
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      Caption         =   "轧辊状态"
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
Attribute VB_Name = "CGF2040C"
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
'-- Program Name      轧辊/轴承座和轴承的库存管理界面
'-- Program ID        CGF2040C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Yang Meng
'-- Coder             Yang Meng
'-- Date              2007.10.31
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
Dim Sc1 As New Collection           'Spread Collection
Dim Sc2 As New Collection           'Spread Collection
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
      Call Gp_Ms_Collection(CBO_ROLL_ID, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(CBO_PLT, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_roll_sts, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(TXT_ROLL_MAKER, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"

    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc2.Add Item:=ss2, Key:="Spread"
    Sc3.Add Item:=ss3, Key:="Spread"

    Proc_Sc.Add Item:=Sc1, Key:="Sc1"
    Proc_Sc.Add Item:=Sc2, Key:="Sc2"
    Proc_Sc.Add Item:=Sc3, Key:="Sc3"

    'Duplicate Count
    iDupCnt = 1

    'Sum Column Count
    iSumCnt = 1

    'Sum Column Setting
    iSumCol.Add Item:=4

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0

End Sub

Private Sub CBO_ROLL_ID_Click()

If Mid(CBO_ROLL_ID.Text, 1, 1) = "J" Or Mid(CBO_ROLL_ID.Text, 1, 1) = "C" Then
   ULabel16.Caption = "轧辊号"
End If
'   Select Case CBO_ROLL_ID.ListIndex
'          Case 0
'               CBO_ROLL_ID.Text = "R"
'               ULabel16.Caption = "轧辊号"
'          Case 1
'               CBO_ROLL_ID.Text = "C"
'               ULabel16.Caption = "轴承座号"
'          Case 2
'               CBO_ROLL_ID.Text = "B"
'               ULabel16.Caption = "轴承号"
'   End Select

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
Dim sQuery_load As String
    Screen.MousePointer = vbHourglass

    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define

    Call Gp_Sp_Setting(Proc_Sc("Sc1")("Spread"), False)
    Call Gp_Sp_Setting(Proc_Sc("Sc2")("Spread"), False)
    Call Gp_Sp_Setting(Proc_Sc("Sc3")("Spread"), False)

    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc1")("Spread"))
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc2")("Spread"))
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc3")("Spread"))

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gf_Sp_Cls(Proc_Sc("Sc1"))
    Call Gf_Sp_Cls(Proc_Sc("Sc2"))
    Call Gf_Sp_Cls(Proc_Sc("Sc3"))

    Call Gp_Sp_ColGet(Proc_Sc("Sc1")("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc2")("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc3")("Spread"), "G-System.INI", Me.Name)

    Screen.MousePointer = vbDefault
    
    CBO_ROLL_ID.Clear
    sQuery_load = "SELECT ROLL_NO FROM GP_ROLL3 ORDER BY SUBSTR(ROLL_NO,1,1) DESC, SUBSTR(ROLL_NO,2,1),SUBSTR(ROLL_NO,6,2)"
    Call Gf_ComboAdd(M_CN1, CBO_ROLL_ID, sQuery_load)


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("Sc1")("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc2")("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc3")("Spread"), "G-System.INI", Me.Name)

    Set rControl = Nothing

    Set Mc1 = Nothing
    Set Sc1 = Nothing
    Set Sc2 = Nothing
    Set Sc3 = Nothing
    Set Proc_Sc = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(Proc_Sc("Sc1")) And Gf_Sp_Cls(Proc_Sc("Sc2")) And Gf_Sp_Cls(Proc_Sc("Sc3")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        ULabel16.Caption = "轧辊号"
    End If

End Sub

Public Sub Form_Exc()

    Call Gp_Sp_Excel(Me, Proc_Sc("Sc1")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc2")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc3")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ref()

    Dim sQuery_R As String
    Dim sQuery_B As String
    Dim sQuery_C As String
    Dim sMesg As String

    sQuery_R = "Select  ROLL_NO ,ROLL_IN_DIA, ROLL_DIA ,PLAN_DIA, ROLL_PRICE,"
    sQuery_R = sQuery_R + " (CASE WHEN ROLL_DIA <= PLAN_DIA THEN 0 ELSE ROUND(ROLL_PRICE * (ROLL_DIA - PLAN_DIA)/(ROLL_IN_DIA - PLAN_DIA),2) END), ROLL_WGT,"
    sQuery_R = sQuery_R + " (CASE WHEN ROLL_DIA <= PLAN_DIA THEN 0 ELSE ROUND(ROLL_WGT * (ROLL_DIA - PLAN_DIA)/(ROLL_IN_DIA - PLAN_DIA),2) END), (ROLL_IN_C_HARD+ROLL_IN_W_HARD+ROLL_IN_D_HARD)/3 ,"
    sQuery_R = sQuery_R + " ROLL_MATERIAL ,ROLL_MAKER_NO ,ROLL_MAKER, ROLL_STATUS , "
    'sQuery_R = sQuery_R + " SUBSTR(ROLL_IN_DATE,1,4) || ' ' || SUBSTR(ROLL_IN_DATE,5,2) ||' ' || SUBSTR(ROLL_IN_DATE,5,2),"
    sQuery_R = sQuery_R + " TO_DATE(ROLL_IN_DATE,'YYYYMMDDHH24MISS'),"
    'sQuery_R = sQuery_R + "SUBSTR(ROLL_DISUSE_DATE,1,4) || ' ' || SUBSTR(ROLL_DISUSE_DATE,5,2) || ' ' || SUBSTR(ROLL_DISUSE_DATE,7,2),"
    sQuery_R = sQuery_R + " TO_DATE(ROLL_DISUSE_DATE,'YYYYMMDDHH24MISS'), "
    sQuery_R = sQuery_R + "ROLL_USE_CNT ,TOT_MILL_WGT ,DECODE((ROLL_IN_DIA-ROLL_DIA),0,0,ROUND(TOT_MILL_WGT/(ROLL_IN_DIA-ROLL_DIA),2)),TOT_MILL_LEN ,GRID_CNT"
    sQuery_R = sQuery_R + "  From GP_ROLL3  "
    sQuery_R = sQuery_R + " Where ROLL_NO  Like '" + Trim(CBO_ROLL_ID.Text) + "%' "
    sQuery_R = sQuery_R + "   AND PLT      Like '" + Trim(CBO_PLT.Text) + "%' "
    sQuery_R = sQuery_R + "   AND ROLL_STATUS      Like '" + Trim(txt_roll_sts.Text) + "%' "
    sQuery_R = sQuery_R + "   AND NVL(ROLL_MAKER,' ')       Like '" + Trim(TXT_ROLL_MAKER.Text) + "%' "
    sQuery_R = sQuery_R + "   Order by SUBSTR(ROLL_NO,1,1) DESC,SUBSTR(ROLL_NO,2,1), SUBSTR(ROLL_NO,6,2) "

    sQuery_B = "Select BEARING_ID ,IN_DIA ,OUT_DIA,WID,STATUS,SUBSTR(IN_WH_DATE,1,8),SUBSTR(DISUSE_DATE,1,8),DISUSE_RES,MAKER"
    sQuery_B = sQuery_B + "  From GP_BEARING3  "
    sQuery_B = sQuery_B + " Where BEARING_ID  Like '" + Trim(CBO_ROLL_ID.Text) + "%' "
    sQuery_B = sQuery_B + " Order by BEARING_ID "

    sQuery_C = "Select CHOCK_ID ,IN_DIA ,OUT_DIA,MATERIAL,STATUS,SUBSTR(IN_WH_DATE,1,8),SUBSTR(DISUSE_DATE,1,8),DISUSE_RES,MAKER"
    sQuery_C = sQuery_C + "  From GP_CHOCK3  "
    sQuery_C = sQuery_C + " Where CHOCK_ID    Like '" + Trim(CBO_ROLL_ID.Text) + "%' "
    sQuery_C = sQuery_C + " Order by CHOCK_ID "

    sMesg = Gf_Ms_NeceCheck(nControl)
    If sMesg = "OK" Then

        sMesg = Gf_Ms_NeceCheck2(mControl)
        If sMesg = "OK" Then

            If Mid(Trim(CBO_ROLL_ID.Text), 1, 1) = "J" Or Mid(Trim(CBO_ROLL_ID.Text), 1, 1) = "C" Then
                 If Gf_Only_Display(M_CN1, Proc_Sc("Sc1"), sQuery_R) Then
                       Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                 End If
            ElseIf Mid(Trim(CBO_ROLL_ID.Text), 1, 1) = "B" Then
                 If Gf_Only_Display(M_CN1, Proc_Sc("Sc3"), sQuery_B) Then
                       Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                 End If
            ElseIf Mid(Trim(CBO_ROLL_ID.Text), 1, 1) = "C" Then
                 If Gf_Only_Display(M_CN1, Proc_Sc("Sc2"), sQuery_C) Then
                       Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                 End If
            Else
                  sMesg = sMesg + " 必须输入 'R' or 'B' or 'C' "
                  Call Gp_MsgBoxDisplay(sMesg)
            End If

        Else
            sMesg = sMesg + " 输入项必须匹配字段长度"
            Call Gp_MsgBoxDisplay(sMesg)
        End If

    Else
        sMesg = sMesg + " 必须输入必要项"
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

Private Sub ss2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss3_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)

    Call Gp_Sp_Sort(Proc_Sc("Sc1")("Spread"), Col, ROW)

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

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal ROW As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If ROW > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub

Private Sub ss2_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal ROW As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If ROW > 0 Then
        Set Active_Spread = Me.ss2
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub

Private Sub ss3_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal ROW As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If ROW > 0 Then
        Set Active_Spread = Me.ss3
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal ROW As Long)

    Call Gp_Sp_Sort(Proc_Sc("Sc2")("Spread"), Col, ROW)

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
End Sub

Private Sub ss3_Click(ByVal Col As Long, ByVal ROW As Long)
 
    Call Gp_Sp_Sort(Proc_Sc("Sc3")("Spread"), Col, ROW)

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
End Sub

Private Sub TXT_ROLL_MAKER_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "G0005"
        DD.rControl.Add Item:=TXT_ROLL_MAKER


        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If
End Sub

Private Sub txt_roll_sts_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "G0007"
        DD.rControl.Add Item:=txt_roll_sts


        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If
End Sub
