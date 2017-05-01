VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form ACA1040C 
   Caption         =   "物料进程现状查询_ACA1040C"
   ClientHeight    =   6825
   ClientLeft      =   885
   ClientTop       =   2265
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6825
   ScaleWidth      =   11865
   WindowState     =   2  'Maximized
   Begin VB.ComboBox combo_ord_item 
      BackColor       =   &H00C0FFFF&
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
      Left            =   2895
      TabIndex        =   1
      Tag             =   "订单号"
      Top             =   90
      Width           =   765
   End
   Begin VB.TextBox text_proc_cd_mate 
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
      Left            =   7305
      TabIndex        =   3
      Top             =   90
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.TextBox text_proc_cd 
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
      Left            =   5775
      MaxLength       =   3
      TabIndex        =   2
      Top             =   90
      Width           =   600
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   330
      Left            =   8565
      Max             =   1
      Min             =   99
      TabIndex        =   6
      Top             =   82
      Value           =   1
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox text_bb_ord_item 
      BackColor       =   &H00C0FFFF&
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
      Left            =   8190
      MaxLength       =   2
      TabIndex        =   4
      Top             =   90
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.TextBox text_bb_ord_no 
      BackColor       =   &H00C0FFFF&
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
      Left            =   1335
      MaxLength       =   11
      TabIndex        =   0
      Tag             =   "订单号"
      Top             =   90
      Width           =   1260
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8715
      Left            =   60
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   525
      Width           =   15120
      _Version        =   393216
      _ExtentX        =   26670
      _ExtentY        =   15372
      _StockProps     =   64
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
      MaxCols         =   11
      MaxRows         =   1
      ProcessTab      =   -1  'True
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "ACA1040C.frx":0000
   End
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   135
      Top             =   90
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "订单号"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   4530
      Top             =   90
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "物料状态"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00404040&
      X1              =   105
      X2              =   15165
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   105
      X2              =   15275
      Y1              =   450
      Y2              =   450
   End
   Begin VB.Line Line2 
      X1              =   2655
      X2              =   2835
      Y1              =   240
      Y2              =   240
   End
End
Attribute VB_Name = "ACA1040C"
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
'-- Program ID        ACA1040C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Yang Zhibin
'-- Date              2003.9.8
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

Public Active_CForm As String       'Form Active

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

Dim iSumCol As New Collection       'Sum Column

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim iCount As Integer

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"
         
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
      Call Gp_Ms_Collection(Text_BB_ORD_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(Combo_ORD_ITEM, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'   Call Gp_Ms_Collection(Text_BB_ORD_ITEM, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(Text_PROC_CD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              
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

Private Sub Combo_ORD_ITEM_LostFocus()

    Dim S As String
    
    If Len(Combo_ORD_ITEM.Text) = 1 Then
       S = Combo_ORD_ITEM.Text
       Combo_ORD_ITEM.Text = "0" + S
    End If
      
End Sub

Private Sub Form_Activate()
    If Active_CForm <> "" Then
        Call Form_Ref
        Active_CForm = ""
    End If
    
    
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

    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)
    
    Set rControl = Nothing
    
    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set Proc_Sc = Nothing
    Set iSumCol = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Combo_ORD_ITEM.Clear
    End If
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ref()

    Dim sProd_cd As String
    Dim sQuery As String
    Dim SMESG As String
    Dim S As String
   

    If Text_BB_ORD_NO.Text = "" Or Combo_ORD_ITEM.Text = "" Then
        Call MsgBox("订单号或序号不能为空!" & Chr(10) & "请输入。", vbExclamation + vbOKOnly, "警告")
    Else
     
        If Combo_ORD_ITEM.Text <> "" Then
           If Len(Combo_ORD_ITEM.Text) = 1 Then
              S = Combo_ORD_ITEM.Text
              Combo_ORD_ITEM.Text = "0" + S
           End If
        End If
        
        sQuery = "SELECT PROD_CD FROM BP_ORDER_ITEM WHERE ORD_NO = '" + Text_BB_ORD_NO.Text + "' AND ORD_ITEM = '" + Combo_ORD_ITEM.Text + "' "
        sProd_cd = Gf_CodeFind(M_CN1, sQuery)
        sQuery = ""
    
        If Mid(Trim(Text_PROC_CD.Text), 1, 1) = "B" Or Trim(Text_PROC_CD.Text) = "" Then
        
            sQuery = " SELECT GF_COMNNAMEFIND('C0004',A.PROC_CD) PROC_CD, A.PROD_CD PROD_CD, A.SLAB_NO SLAB_NO,null  CUR_INV,null LOC, A.SLAB_THK THK, A.SLAB_WID WID, "
            sQuery = sQuery + " A.SLAB_LEN LEN, A.SLAB_WGT WGT1, SUM(B.WGT) WGT2 ,0 WGT3"
            sQuery = sQuery + "   FROM  nisco.EP_SLAB_INS A, nisco.EP_SLAB_DES B "
            sQuery = sQuery + "  WHERE  B.ORD_NO    =     '" + Text_BB_ORD_NO.Text + "' "
            sQuery = sQuery + "    AND  B.ORD_ITEM  =     '" + Combo_ORD_ITEM.Text + "' "
            sQuery = sQuery + "    AND  B.REC_STS   < '3' "
            sQuery = sQuery + "    AND  B.BLOCK_SEQ > '00' "
            sQuery = sQuery + "    AND  B.SEQ       > '00' "
            sQuery = sQuery + "    AND  B.PROC_CD   LIKE  '" + Text_PROC_CD.Text + "%' "
            sQuery = sQuery + "    AND  B.SLAB_NO   = A.SLAB_NO "
            sQuery = sQuery + "  GROUP  BY A.PROC_CD, A.PROD_CD, A.SLAB_NO,A.SLAB_THK, A.SLAB_WID, A.SLAB_LEN, A.SLAB_WGT "
        
        End If
        
        If Trim(sProd_cd) <> "SL" And (Mid(Trim(Text_PROC_CD.Text), 1, 2) = "CA" Or Trim(Text_PROC_CD.Text) = "") Then
            
            If Mid(Trim(Text_PROC_CD.Text), 1, 1) = "B" Or Trim(Text_PROC_CD.Text) = "" Then
                sQuery = sQuery + "  UNION ALL "
            End If
            
            sQuery = sQuery + " SELECT GF_COMNNAMEFIND('C0004',A.PROC_CD) PROC_CD, A.PROD_CD PROD_CD, A.SLAB_NO SLAB_NO,A.CUR_INV, A.LOC, A.THK THK, A.WID WID, "
            sQuery = sQuery + " A.LEN LEN, A.WGT WGT1, SUM(B.WGT) WGT2,NVL(A.LOAD_WGT,0) WGT3"
            sQuery = sQuery + "   FROM  nisco.FP_SLAB A, nisco.FP_SLAB_DES B "
            sQuery = sQuery + "  WHERE  B.ORD_NO    =     '" + Text_BB_ORD_NO.Text + "' "
            sQuery = sQuery + "    AND  B.ORD_ITEM  =     '" + Combo_ORD_ITEM.Text + "' "
            sQuery = sQuery + "    AND  B.REC_STS   < '3' "
            sQuery = sQuery + "    AND  B.BLOCK_SEQ > '00' "
            sQuery = sQuery + "    AND  B.SEQ       > '00' "
            sQuery = sQuery + "    AND  B.PROC_CD   LIKE  '" + Text_PROC_CD.Text + "%' "
            sQuery = sQuery + "    AND  B.SLAB_NO   = A.SLAB_NO "
            sQuery = sQuery + "    AND  A.ORD_FL    = '1' "
            sQuery = sQuery + "  GROUP  BY A.PROC_CD, A.PROD_CD, A.SLAB_NO,A.CUR_INV, A.LOC, A.THK, A.WID, A.LEN, A.WGT ,A.LOAD_WGT"
            
        End If
        
        If (Mid(Trim(Text_PROC_CD.Text), 1, 1) = "B" Or Trim(Text_PROC_CD.Text) = "") Or _
           Trim(sProd_cd) <> "SL" And (Mid(Trim(Text_PROC_CD.Text), 1, 2) = "CA" Or Trim(Text_PROC_CD.Text) = "") Then
            sQuery = sQuery + "  UNION ALL "
        End If
        
        If Trim(sProd_cd) = "SL" Then
        
            sQuery = sQuery + " SELECT GF_COMNNAMEFIND('C0004',A.PROC_CD) PROC_CD, A.PROD_CD PROD_CD, A.SLAB_NO SLAB_NO, A.CUR_INV, A.LOC, A.THK THK, A.WID WID, "
            sQuery = sQuery + "        A.LEN LEN, A.WGT WGT1, A.WGT WGT2,NVL(A.LOAD_WGT,0) WGT3"
            sQuery = sQuery + "   FROM  nisco.FP_SLAB A "
            sQuery = sQuery + "  WHERE  A.ORD_FL    =     '1' "
            sQuery = sQuery + "    AND  A.ORD_NO    =     '" + Text_BB_ORD_NO.Text + "' "
            sQuery = sQuery + "    AND  A.ORD_ITEM  =     '" + Combo_ORD_ITEM.Text + "' "
            sQuery = sQuery + "    AND  A.PROC_CD   LIKE  '" + Text_PROC_CD.Text + "%' "
            
            If Text_PROC_CD.Text = "XAF" Then
                sQuery = sQuery + "    AND  A.REC_STS   = '3' "
            Else
                sQuery = sQuery + "    AND  (A.REC_STS   = '2' OR ( A.REC_STS = '3' AND A.PROC_CD = 'XAF'))"
            End If
            
            sQuery = sQuery + "  ORDER  BY PROC_CD, SLAB_NO "
        
        ElseIf Trim(sProd_cd) = "PP" Then
        
            sQuery = sQuery + " SELECT  GF_COMNNAMEFIND('C0004',A.PROC_CD) PROC_CD, A.PROD_CD PROD_CD, A.PLATE_NO SLAB_NO,  A.CUR_INV, A.LOC, A.THK THK, A.WID WID, A.LEN LEN, "
            sQuery = sQuery + "         A.WGT WGT1, A.WGT WGT2,NVL(A.LOAD_WGT,0) WGT3"
            sQuery = sQuery + "   FROM  nisco.GP_PLATE A "
            sQuery = sQuery + "  WHERE  A.ORD_FL    =     '1' "
            sQuery = sQuery + "    AND  A.PROD_CD   =     'PP' "
            sQuery = sQuery + "    AND  A.ORD_NO    =     '" + Text_BB_ORD_NO.Text + "' "
            sQuery = sQuery + "    AND  A.ORD_ITEM  =     '" + Combo_ORD_ITEM.Text + "' "
            sQuery = sQuery + "    AND  A.PROC_CD   LIKE  '" + Text_PROC_CD.Text + "%' "
            
            If Text_PROC_CD.Text = "XAF" Then
                sQuery = sQuery + "    AND  A.REC_STS   = '3' "
            Else
                sQuery = sQuery + "    AND  (A.REC_STS   <= '2'  OR ( A.REC_STS = '3' AND A.PROC_CD = 'XAF'))"
            End If
            
            sQuery = sQuery + "  ORDER  BY PROC_CD, SLAB_NO "
        
        ElseIf Trim(sProd_cd) = "HC" Then
        
            sQuery = sQuery + " SELECT  GF_COMNNAMEFIND('C0004',A.PROC_CD) PROC_CD, A.PROD_CD PROD_CD, A.COIL_NO SLAB_NO,  A.CUR_INV, A.LOC, A.THK THK, A.WID WID, A.LEN LEN, "
            sQuery = sQuery + "         A.WGT WGT1, A.WGT WGT2,NVL(A.LOAD_WGT,0) WGT3"
            sQuery = sQuery + "   FROM  nisco.GP_COIL A "
            sQuery = sQuery + "  WHERE  A.ORD_FL    =     '1' "
            sQuery = sQuery + "    AND  A.ORD_NO    =     '" + Text_BB_ORD_NO.Text + "' "
            sQuery = sQuery + "    AND  A.ORD_ITEM  =     '" + Combo_ORD_ITEM.Text + "' "
            sQuery = sQuery + "    AND  A.PROC_CD   LIKE  '" + Text_PROC_CD.Text + "%' "
            
            If Text_PROC_CD.Text = "XAF" Then
                sQuery = sQuery + "    AND  A.REC_STS   = '3' "
            Else
                sQuery = sQuery + "    AND  (A.REC_STS   = '2' OR ( A.REC_STS = '3' AND A.PROC_CD = 'XAF'))"
            End If
            
            sQuery = sQuery + "  ORDER  BY PROC_CD, SLAB_NO "
        
        End If
        
        SMESG = Gf_Ms_NeceCheck(nControl)
        If SMESG = "OK" Then
        
            SMESG = Gf_Ms_NeceCheck2(mControl)
            If SMESG = "OK" Then
            
                 If Gf_Only_Display(M_CN1, Proc_Sc("Sc"), sQuery) Then
                    ss1.OperationMode = OperationModeNormal
                    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                 End If
                 
            Else
                SMESG = SMESG + " Must input according to length of item"
                Call Gp_MsgBoxDisplay(SMESG)
            End If
            
        Else
            SMESG = SMESG + " Must input necessarily"
            Call Gp_MsgBoxDisplay(SMESG)
        End If
        
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

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
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

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    If Row > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If
    
End Sub

Private Sub text_bb_ord_no_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sQuery As String
    
    If Len(Trim(Text_BB_ORD_NO.Text)) = Text_BB_ORD_NO.MaxLength Then
    
        If Combo_ORD_ITEM.Text <> "" Then Exit Sub
        
        Text_BB_ORD_NO.Text = StrConv(Text_BB_ORD_NO.Text, vbUpperCase)
        sQuery = " SELECT ORD_ITEM FROM CP_PRC WHERE ORD_NO = '" & Trim(Text_BB_ORD_NO.Text) & "'"
        Call Gf_ComboAdd(M_CN1, Combo_ORD_ITEM, sQuery)
       
       'If Combo_ORD_ITEM.ListCount <> 0 Then
       '   Combo_ORD_ITEM.ListIndex = 0
       'End If
    Else
        Combo_ORD_ITEM.Clear
    End If

End Sub

Private Sub Text_PROC_CD_Change()

    If Not Text_PROC_CD.Text = "" Then
    
        If Len(Text_PROC_CD.Text) = Text_PROC_CD.MaxLength Then
        
            Text_PROC_CD.Text = StrConv(Text_PROC_CD.Text, vbUpperCase)
            Select Case Text_PROC_CD.Text
               Case "BAA", "BAB", "BAC", "BAD", "BAE", "BAF"
               Case "BBA", "BBB", "BBC", "BBD", "BBE", "BBF"
               Case "BCA", "BCB", "BCC", "BCD", "BCE", "BCF"
               Case "BDA", "BDB", "BDC", "BDD", "BDE", "BDF"
               Case "BEA", "BEB", "BEC", "BED", "BEE", "BEF"
               Case "BFA", "BFB", "BFC", "BFD", "BFE", "BFF"
               Case "CAA", "CAB", "CAC", "CAD", "CAE", "CAF"
               Case "CBA", "CBB", "CBC", "CBD", "CBE", "CBF"
               Case "CGA", "CGB", "CGC", "CGD", "CGE", "CGF"
               Case "DAA", "DAB", "DAC", "DAD", "DAE", "DAF"
               Case "DZB", "DZE"
               Case "DBA", "DBB", "DBC", "DBD", "DBE", "DBF"
               Case "QAA", "QAB", "QAC", "QAD", "QAE", "QAF"
               Case "XAA", "XAB", "XAC", "XAD", "XAE", "XAF"
               Case ""
            Case Else
                  Call MsgBox("进程代码不符合规范!" & Chr(10) & "请更正。", vbExclamation + vbOKOnly, "警告")
                  Text_PROC_CD.Text = ""
                  'Text_PROC_CD_Name.Text = ""
            End Select
            
        End If
        
    End If

End Sub

Private Sub Text_PROC_CD_DblClick()

    Call Text_PROC_CD_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub Text_PROC_CD_KeyUp(KeyCode As Integer, Shift As Integer)
   
   If KeyCode = vbKeyF4 Then
 
        DD.sWitch = "MS"
        DD.sKey = "C0004"

        DD.rControl.Add Item:=Text_PROC_CD
        DD.rControl.Add Item:=Text_PROC_CD_mate
   
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
        
    End If

    If Len(Trim(Text_PROC_CD.Text)) = Text_PROC_CD.MaxLength Then
        Text_PROC_CD_mate.Text = Gf_ComnNameFind(M_CN1, "C0004", Text_PROC_CD.Text, 2)
    Else
        Text_PROC_CD_mate.Text = ""
    End If

End Sub
