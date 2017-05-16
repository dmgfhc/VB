VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AQA0070C 
   Caption         =   "客户特殊要求共用信息查询_AQA0070C"
   ClientHeight    =   9090
   ClientLeft      =   30
   ClientTop       =   1920
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_CUST_NAME 
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
      Left            =   2580
      TabIndex        =   2
      Top             =   120
      Width           =   4335
   End
   Begin VB.TextBox txt_CUST_CD 
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
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   1245
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8625
      Left            =   105
      TabIndex        =   0
      Top             =   525
      Width           =   15105
      _Version        =   393216
      _ExtentX        =   26644
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
      MaxCols         =   43
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQA0070C.frx":0000
   End
   Begin InDate.ULabel ULabel12 
      Height          =   315
      Index           =   1
      Left            =   150
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "客户"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin Threed.SSCommand Com_STD_CHEM 
      Height          =   360
      Left            =   11850
      TabIndex        =   3
      Top             =   75
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   635
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "客户特殊要求成分"
      BevelWidth      =   1
   End
   Begin Threed.SSCommand Com_STD_MATR 
      Height          =   360
      Left            =   13530
      TabIndex        =   4
      Top             =   75
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   635
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "客户特殊要求材质"
      BevelWidth      =   1
   End
   Begin Threed.SSCommand Com_STD_CHEM_MAT 
      Height          =   360
      Left            =   10140
      TabIndex        =   5
      Top             =   75
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   635
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "客户特殊要求详细"
      BevelWidth      =   1
   End
End
Attribute VB_Name = "AQA0070C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       质量管理
'-- Sub_System Name   质量标准管理
'-- Program Name      客户特殊要求共用信息查询
'-- Program ID        AQA0070C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Lee Qing Yu
'-- Coder             Lee Qing Yu
'-- Date              2003.5.19
'-- Description       客户特殊要求共用信息查询
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
Dim Sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "PopSheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
      Call Gp_Ms_Collection(txt_CUST_CD, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_CUST_NAME, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   
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
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 35, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 36, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 37, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 38, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 39, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 40, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 41, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 42, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 43, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AQA0070C.P_SREFER", Key:="P-R"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc"
     
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
         
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Code Name Find --------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo Err_Track:
    
    Dim oCodeName As Object
    Dim sCode As String
        
    Select Case Me.ActiveControl.Name
        Case "txt_CUST_CD"
            sCode = "CUST_CD"
            Set oCodeName = txt_CUST_NAME
    End Select
    
    If sCode = "" Then Exit Sub
    
    Call Gp_MS_CodeNameFind(KeyCode, sCode, Me.ActiveControl, oCodeName)
    
    Set oCodeName = Nothing
Err_Track:
End Sub


Private Sub Com_STD_CHEM_MAT_Click()
 With ss1
    .Row = .ActiveRow
    .Col = 1
    AQA0110C.txt_CUST_SPEC_NO.Text = .Text
    
    Call AQA0110C.Form_Ref
    
 End With
 AQA0110C.Show
 AQA0110C.SetFocus

End Sub

Private Sub Com_STD_CHEM_Click()
 With ss1
    .Row = .ActiveRow
    .Col = 1
    AQA0090C.txt_CUST_SPEC_NO.Text = .Text
    .Col = 2
    AQA0090C.txt_CUST_STD_NAME.Text = .Text
    .Col = 23
    AQA0090C.uLab_CUST_DETAIL.Caption = .Text
    
 End With
 AQA0090C.Show
 AQA0090C.SetFocus
 
End Sub

Private Sub Com_STD_MATR_Click()
 With ss1
    .Row = .ActiveRow
    .Col = 1
    AQA0100C.txt_CUST_SPEC_NO.Text = .Text
    .Col = 23
    AQA0100C.ulb_Detail.Caption = .Text
 End With
 AQA0100C.Show
 AQA0100C.SetFocus

End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    Call GP_MENU_SHOW_HIDE("04F11F12F")
    MDIMain.MenuTool.Buttons(7).Enabled = True
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name, True)
    
    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
    Call GP_MENU_SHOW_HIDE("04F11F12F")
    MDIMain.MenuTool.Buttons(7).Enabled = True
'    .Buttons(7).Enabled = False                 'Row Insert
'                 .Buttons(8).Enabled = False                 'Row Delete
'                 .Buttons(9).Enabled = False                 'Row Cancel
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

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
    Set Sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        Call GP_MENU_SHOW_HIDE("04F11F12F")
        rControl(1).SetFocus
    End If

End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    Dim sMesg As String
    

        
            If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                Call GP_SELECT_ROW(ss1, 1)
                Call GP_MENU_SHOW_HIDE("04F11F12F")
                MDIMain.MenuTool.Buttons(7).Enabled = True
                
                Exit Sub
            End If
            

    
    Exit Sub

Refer_Err:

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

'Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)
'
'    'Spread --> Control Value Move
'
'    Dim iRow As Long
'
'    With ss1
'
'        iRow = Row
'
'       If Row <> 0 Then
'
'           Load AQA0080C
'
'           .Row = Row
'
'           .Col = 1: AQA0080C.txt_CUST_CD.Text = Left(.Text, 6)
'           .Col = 1: AQA0080C.txt_CUST_SQ.Text = Right(.Text, 3)
'           .Col = 2: AQA0080C.txt_CUST_NAME.Text = .Text
'           .Col = 3: AQA0080C.txt_PROD_CD.Text = .Text
'           .Col = 4: AQA0080C.txt_PROD_NAME.Text = .Text
'           .Col = 5: AQA0080C.txt_STEEL_GRD.Text = .Text
'           .Col = 6: AQA0080C.txt_STEEL_GRD_Name.Text = .Text
'           .Col = 7: AQA0080C.txt_ENDUSE_CD.Text = .Text
'           .Col = 8: AQA0080C.txt_ENDUSE_NAME.Text = .Text
'           .Col = 9: AQA0080C.txt_STDSPEC.Text = .Text
'           .Col = 10: AQA0080C.sdb_STDSPEC_YY.Text = .Text
'           .Col = 11: AQA0080C.txt_DEV_STD_CD.Text = .Text
'           .Col = 12: AQA0080C.txt_Nisco_Quality_No.Text = .Text
'           .Col = 13: AQA0080C.txt_MLT_STD_NO.Text = .Text
'           .Col = 14: AQA0080C.txt_MILL_STD_NO.Text = .Text
'
'           .Col = 15: AQA0080C.txt_HTM_SHOT_BLAST.Text = .Text
'           .Col = 16: AQA0080C.txt_HTM_SHOT_BLAST_NAME.Text = .Text
'           .Col = 17: AQA0080C.txt_HTM_METH_CD_1.Text = .Text
'           .Col = 18: AQA0080C.txt_HTM_METH_NAME_1.Text = .Text
'           .Col = 19: AQA0080C.txt_HTM_COND_CD_1.Text = .Text
'           .Col = 20: AQA0080C.txt_HTM_COND_NAME_1.Text = .Text
'           .Col = 21: AQA0080C.txt_HTM_METH_CD_2.Text = .Text
'           .Col = 22: AQA0080C.txt_HTM_METH_NAME_2.Text = .Text
'           .Col = 23: AQA0080C.txt_HTM_COND_CD_2.Text = .Text
'           .Col = 24: AQA0080C.txt_HTM_COND_NAME_2.Text = .Text
'           .Col = 25: AQA0080C.txt_HTM_METH_CD_3.Text = .Text
'           .Col = 26: AQA0080C.txt_HTM_METH_NAME_3.Text = .Text
'           .Col = 27: AQA0080C.txt_HTM_COND_CD_3.Text = .Text
'           .Col = 28: AQA0080C.txt_HTM_COND_NAME_3.Text = .Text
'
'           .Col = 29: AQA0080C.txt_DRT_CNF_TYP.Text = .Text
'           .Col = 31: AQA0080C.sdb_THK_MIN.Text = .Text
'           .Col = 32: AQA0080C.sdb_THK_MAX.Text = .Text
'           .Col = 33: AQA0080C.sdb_WID_MIN.Text = .Text
'           .Col = 34: AQA0080C.sdb_WID_MAX.Text = .Text
'           .Col = 35: AQA0080C.sdb_LEN_MIN.Text = .Text
'           .Col = 36: AQA0080C.sdb_LEN_MAX.Text = .Text
'           .Col = 37: AQA0080C.txt_CUST_SPEC_DETAIL.Text = .Text
'           .Col = 38: AQA0080C.txt_INS_DATE.Text = .Text
'           .Col = 39: AQA0080C.txt_ins_emp.Text = .Text
'           .Col = 40: AQA0080C.txt_ins_name.Text = .Text
'           .Col = 41: AQA0080C.txt_UPD_DATE.Text = .Text
'           .Col = 42: AQA0080C.txt_UPD_EMP.Text = .Text
'           .Col = 43: AQA0080C.txt_upd_name.Text = .Text
'
'       End If
'
'        AQA0080C.Show 1
'
'   End With
'
'   Call GP_MENU_SHOW_HIDE("04F11F12F")
'   Call GP_SELECT_ROW(ss1, iRow)
'   MDIMain.MenuTool.Buttons(7).Enabled = True
'
'End Sub

Private Sub ss1_LeaveRow(ByVal Row As Long, ByVal RowWasLast As Boolean, ByVal RowChanged As Boolean, ByVal AllCellsHaveData As Boolean, ByVal NewRow As Long, ByVal NewRowIsLast As Long, Cancel As Boolean)
'    Call GP_SetRowHeaderClear(ss1, NewRow)
End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

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

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub


Private Sub txt_CUST_CD_Change()
    If Trim(txt_CUST_CD.Text) = "" Then
        txt_CUST_NAME.Text = ""
    End If
End Sub
Public Sub Form_Ins()

    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    Call GP_SELECT_ROW(ss1, ss1.ActiveRow)
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 39)
    Call ss1_DblClick(1, ss1.ActiveRow)
End Sub
