VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form ACF2080C 
   Caption         =   "炼钢生产月报查询"
   ClientHeight    =   9375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9375
   ScaleWidth      =   15240
   Tag             =   "工厂"
   WindowState     =   2  'Maximized
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
      Height          =   310
      Left            =   1590
      TabIndex        =   2
      Top             =   135
      Width           =   1770
   End
   Begin VB.TextBox txt_PLT 
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
      Left            =   1035
      MaxLength       =   2
      TabIndex        =   1
      Top             =   135
      Width           =   555
   End
   Begin VB.TextBox txt_PRC_LINE 
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
      Left            =   4725
      MaxLength       =   1
      TabIndex        =   0
      Tag             =   "机号"
      Top             =   135
      Width           =   435
   End
   Begin FPSpread.vaSpread SS1 
      Height          =   8580
      Left            =   180
      TabIndex        =   3
      Top             =   585
      Width           =   14925
      _Version        =   393216
      _ExtentX        =   26326
      _ExtentY        =   15134
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
      MaxCols         =   10
      MaxRows         =   497
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "ACF2080C.frx":0000
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   5850
      Top             =   135
      Width           =   795
      _ExtentX        =   1402
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
   Begin InDate.UDate cbo_date 
      Height          =   315
      Left            =   6705
      TabIndex        =   4
      Tag             =   "日期"
      Top             =   135
      Width           =   1455
      _ExtentX        =   2566
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
      MaxLength       =   6
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   3870
      Top             =   135
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   556
      Caption         =   "机号"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
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
      Left            =   180
      Top             =   135
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   556
      Caption         =   "工厂"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
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
Attribute VB_Name = "ACF2080C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       C System
'-- Sub_System Name   C
'-- Program Name
'-- Program ID
'-- Document No       Q-00-0010(Specification)
'-- Designer          JIA NING
'-- Coder             JIA NING
'-- Date              2003.7.19
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
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim iSumCol As New Collection       'Sum Column

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"
         
         Call Gp_Ms_Collection(txt_PLT, "p", "n", " ", " ", "r", "a", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_PLT_NAME, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_PRC_LINE, "p", "n", " ", " ", "r", "a", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(cbo_date, "p", "n", " ", " ", "r", "a", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   
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
    Sc1.Add Item:=SS1, Key:="Spread"
    
    Proc_Sc.Add Item:=Sc1, Key:="Sc"
    
    'Duplicate Count
    iDupCnt = 1
    
    'Sum Column Count
    iSumCnt = 2
    
    'Sum Column Setting
    iSumCol.Add Item:=8
    iSumCol.Add Item:=9
    
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
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
    
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))
   
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "Z-System.INI", Me.Name)

    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "Z-System.INI", Me.Name)
    
    Set rControl = Nothing
    
    Set Mc1 = Nothing
    Set Sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    End If
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ref()
Dim YEARMONTH As String
YEARMONTH = Mid(cbo_date.Text, 1, 4) + Mid(cbo_date.Text, 6, 2)

'MsgBox "11"

    Dim sQuery As String
    Dim sMesg As String

sQuery = "SELECT A.STLGRD,A.THK_GRP,A.WID_GRP,C.FR_THK,C.TO_THK,D.FR_WID,D.TO_WID,A.PLAN_VALUE,B.B,B.B/A.PLAN_VALUE*100"
sQuery = sQuery + " From"
sQuery = sQuery + "(select STLGRD,"
sQuery = sQuery + " THK_GRP,"
sQuery = sQuery + " WID_GRP,"
sQuery = sQuery + "SUM(WGT) B"
sQuery = sQuery + " From FP_SLAB"
sQuery = sQuery + " WHERE PLT = '" + Trim(txt_PLT.Text) + "'"
sQuery = sQuery + " AND   PRC_LINE = '" + Trim(Me.txt_PRC_LINE.Text) + "'"
 
sQuery = sQuery + "AND   PROD_DATE LIKE '" + Trim(YEARMONTH) + "%'"
sQuery = sQuery + " AND   IN_PLT_CD <>'4'"
sQuery = sQuery + " AND   IN_PLT_CD <>'5'"
sQuery = sQuery + " GROUP BY THK_GRP,WID_GRP,STLGRD "
sQuery = sQuery + " ORDER BY THK_GRP ASC,WID_GRP ASC,STLGRD ASC) B,AP_PROD_PLAN A ,BP_THICK_GRP C,BP_WIDTH_GRP D"
sQuery = sQuery + " Where A.STLGRD = B.STLGRD"
sQuery = sQuery + " AND   A.THK_GRP = B.THK_GRP"
sQuery = sQuery + " AND   A.WID_GRP = B.WID_GRP"
sQuery = sQuery + " AND   A.PLT = '" + Trim(txt_PLT.Text) + "'"
sQuery = sQuery + " AND   A.PRC_LINE ='" + Trim(txt_PRC_LINE.Text) + "'"
sQuery = sQuery + " AND   A.PRC = 'BF'"
sQuery = sQuery + " AND   A.YEAR_MONTH = '" + Trim(YEARMONTH) + "'"
sQuery = sQuery + " AND   A.PROD_CD = 'SL'"
sQuery = sQuery + " AND   A.WID_GRP = D.WID_CD"
sQuery = sQuery + " AND   A.THK_GRP = C.THK_CD"
sQuery = sQuery + " AND   A.PROD_CD = C.PROD_CD"
sQuery = sQuery + " AND   A.PROD_CD = D.PROD_CD"
sQuery = sQuery + " ORDER BY B.STLGRD ASC,B.THK_GRP ASC,B.WID_GRP ASC"


'Exit Sub



    sMesg = Gf_Ms_NeceCheck(nControl)
    
    If sMesg = "OK" Then
    
        sMesg = Gf_Ms_NeceCheck2(mControl)
        If sMesg = "OK" Then

            'If Gf_Only_Display(M_CN1, Proc_Sc("Sc"), sQuery) Then
            'If Gf_Total_Display(M_CN1, Proc_Sc("Sc"), sQuery, iDupCnt, iSumCnt, iSumCol) Then
            'If Gf_Only_Display(M_CN1, Proc_Sc("Sc"), sQuery, iDupCnt, iSumCnt, iSumCol) Then
            If Gf_Stotal_Display(M_CN1, Proc_Sc("Sc"), sQuery, iDupCnt, iSumCnt, iSumCol) Then
            
       
        
            
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            End If
            
        Else
            sMesg = sMesg + " Must input according to length of item"
            Call Gp_MsgBoxDisplay(sMesg)
        End If
    
    Else
        sMesg = sMesg + " Must input necessarily"
        Call Gp_MsgBoxDisplay(sMesg)
    End If
 If SS1.MaxRows <> 0 Then
 Dim I As Integer
 Dim PLANWGT As Double
 Dim TRUEWGT As Double
 
 With SS1
     For I = 1 To .MaxRows
     .Col = 1
     .Row = I
         If .Text = "SUB TOTAL" Or .Text = "TOTAL" Then
            .Col = 8
            PLANWGT = .Value
            .Col = 9
            TRUEWGT = .Value
            .Col = 10
                If PLANWGT = 0 Then
                .Value = 0
                Else
                .Value = TRUEWGT / PLANWGT
                .Value = .Value * 100
                End If
                
        End If
    Next
End With

 
 
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

   ' Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
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
 MDIMain.Mnu_Sorting = False
   If Row > 0 Then
       Set Active_Spread = Me.SS1
        PopupMenu MDIMain.PopUp_Spread
    End If
  
   
End Sub

Private Sub txt_PLT_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0001"
        DD.rControl.Add Item:=txt_PLT
        DD.rControl.Add Item:=txt_PLT_NAME
        
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
        Exit Sub
        
    End If

    If Len(Trim(txt_PLT.Text)) = txt_PLT.MaxLength Then
        txt_PLT_NAME.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_PLT.Text), 2)
    Else
        txt_PLT_NAME.Text = ""
        
    End If



End Sub
