VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form AFL2080C 
   Caption         =   "¿â´æ°åÅ÷ÖÖÀà²éÑ¯½çÃæ_AFL2080C"
   ClientHeight    =   9225
   ClientLeft      =   405
   ClientTop       =   1680
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15225
   WindowState     =   2  'Maximized
   Begin FPSpread.vaSpread ss1 
      Height          =   9090
      Left            =   75
      TabIndex        =   0
      Top             =   60
      Width           =   15105
      _Version        =   393216
      _ExtentX        =   26644
      _ExtentY        =   16034
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ËÎÌå"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   8
      MaxRows         =   2
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AFL2080C.frx":0000
   End
End
Attribute VB_Name = "AFL2080C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       NISCO Production Management System
'-- Sub_System Name
'-- Program Name
'-- Program ID        AFL2080C
'-- Designer          KIM SUNG HO
'-- Coder             KIM SUNG HO
'-- Date              2003.11.11
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

Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"
    
    'Duplicate Col Count
    iDupCnt = 1

    sc1.Add Item:=ss1, Key:="Spread"
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
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
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "F-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "F-System.INI", Me.Name)
    
    Set sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    End If
    
End Sub

Public Sub Form_Ref()

    Dim sQuery As String
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    sQuery = "          SELECT /*+INDEX (A FI_SLAB24)*/ 'ÖÐ¼ä¶©µ¥°åÅ÷' ITEM, A.SLAB_NO, A.STLGRD, GF_STLGRD_DETAIL(A.STLGRD) STLGRD_DETAIL, A.LEN, A.WID, A.THK, A.WGT "
    sQuery = sQuery + "   FROM FP_SLAB A WHERE A.PROD_CD IN ('HC','PP') AND A.REC_STS = '2' AND A.LOC IS NOT NULL "
    
    sQuery = sQuery + " UNION  ALL "
    
    sQuery = sQuery + " SELECT /*+INDEX (A FI_SLAB24)*/ '¶©µ¥°åÅ÷' ITEM, A.SLAB_NO, A.STLGRD, GF_STLGRD_DETAIL(A.STLGRD) STLGRD_DETAIL, A.LEN, A.WID, A.THK, A.WGT "
    sQuery = sQuery + "   FROM FP_SLAB A WHERE A.PROD_CD = 'SL' AND A.REC_STS = '2' AND A.LOC IS NOT NULL "
    
    sQuery = sQuery + " UNION  ALL "
    
    sQuery = sQuery + " SELECT /*+INDEX (A FI_SLAB24)*/ 'Óà²Ä' ITEM, A.SLAB_NO, A.STLGRD, GF_STLGRD_DETAIL(A.STLGRD) STLGRD_DETAIL, A.LEN, A.WID, A.THK, A.WGT "
    sQuery = sQuery + "   FROM FP_SLAB A WHERE A.ORD_FL = '2' AND A.REC_STS = '2' AND A.LOC IS NOT NULL "
    
    sQuery = sQuery + " UNION  ALL "
    
    sQuery = sQuery + " SELECT /*+INDEX (A FI_SLAB24)*/ 'ÐÞÄ¥°åÅ÷' ITEM, A.SLAB_NO, A.STLGRD, GF_STLGRD_DETAIL(A.STLGRD) STLGRD_DETAIL, A.LEN, A.WID, A.THK, A.WGT "
    sQuery = sQuery + "   FROM FP_SLAB A WHERE A.SCR_ORNOT = 'Y' AND A.REC_STS = '2' AND A.LOC IS NOT NULL "
    
    sQuery = sQuery + " UNION  ALL "
    
    sQuery = sQuery + " SELECT /*+INDEX (A FI_SLAB24)*/ 'ÂòÈë°åÅ÷' ITEM, A.SLAB_NO, A.STLGRD, GF_STLGRD_DETAIL(A.STLGRD) STLGRD_DETAIL, A.LEN, A.WID, A.THK, A.WGT "
    sQuery = sQuery + "   FROM FP_SLAB A WHERE A.IN_PLT_CD IN ('4','5') AND A.REC_STS = '2' AND A.LOC IS NOT NULL "
            
    If Gf_Only_Display(M_CN1, Proc_Sc("Sc"), sQuery, iDupCnt) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
            
End Sub

Public Sub Form_Pro()

End Sub

Public Sub Form_Ins()
    
End Sub

Public Sub Spread_Cpy()

End Sub

Public Sub Spread_Pst()

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
