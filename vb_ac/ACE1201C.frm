VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form ACE1201C 
   Caption         =   "相关母板/产品信息_ACE1201C"
   ClientHeight    =   6285
   ClientLeft      =   570
   ClientTop       =   2655
   ClientWidth     =   11115
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   11115
   Begin VB.TextBox TXT_SLAB_NO 
      Enabled         =   0   'False
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
      Left            =   1260
      MaxLength       =   11
      TabIndex        =   0
      Tag             =   "钢种"
      Top             =   135
      Width           =   1275
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   90
      Top             =   135
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      Caption         =   "板 坯 号"
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
   End
   Begin FPSpread.vaSpread SS1 
      Height          =   5595
      Left            =   45
      TabIndex        =   1
      Top             =   630
      Width           =   10995
      _Version        =   393216
      _ExtentX        =   19394
      _ExtentY        =   9869
      _StockProps     =   64
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   12
      MaxRows         =   2
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "ACE1201C.frx":0000
   End
   Begin Threed.SSCommand CMD1 
      Height          =   420
      Left            =   9990
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   90
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   741
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "退出"
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   45
      X2              =   11025
      Y1              =   540
      Y2              =   540
   End
End
Attribute VB_Name = "ACE1201C"
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
'-- Program ID        ACE1201C
'-- Document No       Q-00-0010(Specification)
'-- Designer          jianing
'-- Coder             jianing
'-- Date              2003.9.17
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------

Public FormType   As String           'Form Type
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

Dim iSumCol As New Collection       'Sum Column

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"
         
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
      Call Gp_Ms_Collection(txt_slab_no, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
         
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    'Duplicate Count
    iDupCnt = 1
    
    'Sum Column Count
    iSumCnt = 1
    
    'Sum Column Setting
    iSumCol.Add Item:=4
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    Me.Top = 3000
    Me.Left = 2300

End Sub

Private Sub CMD1_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    
    'Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    Call Form_Ref
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
   
    'Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)

    Screen.MousePointer = vbDefault
    Me.txt_slab_no = ACE1200C.P_SLAB_NO
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)
    
    Set rControl = Nothing
    
    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set Proc_Sc = Nothing
    
   ' Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

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

    Dim sQuery As String
    Dim SMESG As String

    sQuery = " SELECT B.BLOCK_SEQ"
    sQuery = sQuery + " , B.SEQ    "
    sQuery = sQuery + " , B.ORD_NO"
    sQuery = sQuery + " , B.ORD_ITEM"
    sQuery = sQuery + " , B.OVER_FL"
    sQuery = sQuery + " , B.THK"
    sQuery = sQuery + " , B.WID"
    sQuery = sQuery + " , B.LEN"
    sQuery = sQuery + " , B.WGT"
    sQuery = sQuery + " , B.CR_CD"
    sQuery = sQuery + " , B.UST_FL"
    sQuery = sQuery + " , B.TRIM_FL"
    sQuery = sQuery + " FROM  CP_REP_SLAB_D B"
    sQuery = sQuery + " WHERE B.SLAB_NO = '"
    sQuery = sQuery + Trim(txt_slab_no.Text) + "'"
    sQuery = sQuery + "ORDER BY BLOCK_SEQ ASC,SEQ ASC"
    
    SMESG = Gf_Ms_NeceCheck(nControl)
    If SMESG = "OK" Then
    
        SMESG = Gf_Ms_NeceCheck2(mControl)
        If SMESG = "OK" Then

            If Gf_Only_Display(M_CN1, Proc_Sc("Sc"), sQuery) Then
                'Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            End If
            
        Else
            SMESG = SMESG + " Must input according to length of item"
            Call Gp_MsgBoxDisplay(SMESG)
        End If
    
    Else
        SMESG = SMESG + " Must input necessarily"
        Call Gp_MsgBoxDisplay(SMESG)
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

Private Sub Form_Resize()

    If Me.Width - 1320 < 0 Or Me.Height - 1440 < 0 Then
        Me.Width = 1320
        Me.Height = 1440
        Exit Sub
    End If

    ss1.Width = Me.Width - 225
    ss1.Height = Me.Height - 1440
    Line1.X2 = ss1.Width
    CMD1.Left = Me.Width - 1320

End Sub
