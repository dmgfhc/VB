VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AQE2010C 
   Caption         =   "轧后板坯退废表_AQE2010C"
   ClientHeight    =   9135
   ClientLeft      =   1500
   ClientTop       =   2025
   ClientWidth     =   10050
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9135
   ScaleWidth      =   10050
   WindowState     =   2  'Maximized
   Begin VB.TextBox TXT_BOF_CC 
      Height          =   960
      Left            =   4170
      TabIndex        =   11
      Top             =   3105
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.ComboBox cbo_PLT 
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
      ItemData        =   "AQE2010C.frx":0000
      Left            =   9630
      List            =   "AQE2010C.frx":000A
      TabIndex        =   10
      Top             =   540
      Width           =   1485
   End
   Begin VB.TextBox txt_STLGRD_DETAIL 
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
      Height          =   300
      Left            =   5670
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   540
      Width           =   1485
   End
   Begin VB.TextBox txt_STLGRD 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4185
      MaxLength       =   11
      TabIndex        =   8
      Tag             =   "钢种"
      Top             =   540
      Width           =   1485
   End
   Begin VB.ComboBox CBO_GROUP 
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
      ItemData        =   "AQE2010C.frx":0016
      Left            =   9630
      List            =   "AQE2010C.frx":0026
      TabIndex        =   7
      Top             =   135
      Width           =   1485
   End
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   1815
      Begin VB.OptionButton opt_bof 
         BackColor       =   &H00E0E0E0&
         Caption         =   "连铸"
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Value           =   -1  'True
         Width           =   810
      End
      Begin VB.OptionButton opt_CCM 
         BackColor       =   &H00E0E0E0&
         Caption         =   "轧制"
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   120
         Width           =   780
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   30
      Left            =   0
      TabIndex        =   0
      Top             =   9105
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   53
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8220
      Left            =   90
      TabIndex        =   4
      Top             =   960
      Width           =   15105
      _Version        =   393216
      _ExtentX        =   26644
      _ExtentY        =   14499
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   13
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQE2010C.frx":0036
   End
   Begin InDate.ULabel ULabel12 
      Height          =   300
      Index           =   2
      Left            =   2880
      Top             =   135
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      Caption         =   "日期"
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
   Begin InDate.UDate dtp_DATE_FROM 
      Height          =   300
      Left            =   4185
      TabIndex        =   5
      Tag             =   "订单确认日期"
      Top             =   135
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      BackColor       =   16777215
   End
   Begin InDate.UDate dtp_DATE_TO 
      Height          =   300
      Left            =   5670
      TabIndex        =   6
      Tag             =   "订单确认日期"
      Top             =   135
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      BackColor       =   16777215
   End
   Begin InDate.ULabel ULabel8 
      Height          =   300
      Left            =   8325
      Top             =   135
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      Caption         =   "班别"
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
   Begin InDate.ULabel ULabel3 
      Height          =   300
      Index           =   1
      Left            =   2880
      Top             =   540
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      Caption         =   "钢种"
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
      Height          =   300
      Index           =   0
      Left            =   8325
      Top             =   540
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      Caption         =   "轧制工厂"
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
End
Attribute VB_Name = "AQE2010C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       Slab Quality Analysis and Stat.
'-- Sub_System Name
'-- Program Name
'-- Program ID        AQE1050C
'-- Document No       Q-00-0010(Specification)
'-- Designer          HJD
'-- Coder             HJD
'-- Date              2006.09.07
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
Dim Sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim SumCnt1 As Integer
Dim SumCol1 As New Collection


Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()

    Dim i As Long
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"

       'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
           Call Gp_Ms_Collection(TXT_BOF_CC, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(dtp_DATE_FROM, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(dtp_DATE_TO, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(cbo_PLT, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_STLGRD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_STLGRD_DETAIL, " ", " ", " ", " ", " ", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(CBO_GROUP, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------  BELOW EDIT ---------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
    
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
    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    'Sc1.Add Item:="AQE1050C.P_MODIFY", Key:="P-M"
    Sc1.Add Item:="AQE2010C.P_REFER", Key:="P-R"
    'Sc1.Add Item:="AQE1050C.P_ONEROW", Key:="P-O"
    
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------  EDIT  End      ---------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
    
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
        


    'Duplicate Count
    iDupCnt = 1
    
    'Sum Column Count
    SumCnt1 = 7
    
    'Sum Column Setting
    SumCol1.Add Item:=3
    SumCol1.Add Item:=4
    
    For i = 9 To 13
        SumCol1.Add Item:=i
    Next i
    
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
    
    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault
    TXT_BOF_CC = "1"

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
    Set Sc1 = Nothing
    Set Proc_Sc = Nothing
   Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub


Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        'rControl(1).SetFocus
    End If

End Sub

Public Sub Form_Ref()


    Dim irow                  As Integer
    Dim iSLAB_WGT             As Double
    Dim iSLAB_WGT_SUM         As Double
    Dim iJZ                   As Double
    Dim iZXLW                 As Double
    Dim iQTLW                 As Double
    Dim iQP                   As Double
    Dim iQT                   As Double
    Dim i, j                  As Integer
    
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
            
'    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
'
'        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'
'    End If
   Call Gf_Multi_Stotal_Display(M_CN1, Proc_Sc("Sc"), Gf_Ms_MakeQuery(Proc_Sc("Sc").Item("P-R"), "R", Mc1("pControl")), 1, 1, SumCnt1, SumCol1, False)
   Call Sp_Zero_Clear
   
End Sub

'Public Sub Form_Pro()
'
'    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
'        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'    End If
'
'End Sub

'Public Sub Spread_ColumnsSort()
'
'    Spread_ColSort.Show 1
'
'End Sub

Public Sub Spread_Forzens_Setting()

    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol
    
End Sub

Public Sub Spread_Forzens_Cancel()

    Me.ActiveControl.ColsFrozen = 0
    
End Sub

Public Sub Form_Exc()
        
    Call Gp_Sp_Excel(Me, ss1, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Private Sub opt_bof_Click()
    TXT_BOF_CC.Text = "1"
    Call Form_Ref
End Sub

Private Sub opt_ccm_Click()
    TXT_BOF_CC.Text = "2"
    Call Form_Ref
End Sub

Private Sub txt_STLGRD_Change()
    Dim sQuery As String
    If Len(Trim(txt_STLGRD.Text)) = 11 Then
        sQuery = "SELECT STEEL_GRD_DETAIL FROM QP_NISCO_CHMC WHERE STLGRD = '" + Trim(txt_STLGRD.Text) + "'"
        txt_STLGRD_DETAIL.Text = Gf_FloatFind(M_CN1, sQuery)
    Else
        txt_STLGRD_DETAIL.Text = ""
    End If
End Sub
'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Code Name Find --------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Track:
    Dim oCodeName As Object
    Dim sCode As String
    
    Select Case Me.ActiveControl.Name
        
        Case "txt_STLGRD"
            sCode = "STLGRD"
            Set oCodeName = txt_STLGRD_DETAIL
    End Select
    
    If sCode = "" Then Exit Sub
    
    Call Gp_MS_CodeNameFind(KeyCode, sCode, Me.ActiveControl, oCodeName)
    
    Set oCodeName = Nothing
Err_Track:
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    'Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub


'Private Sub ss1_LostFocus()
'
'    lBlkcol1 = 0
'    lBlkcol2 = 0
'    lBlkrow1 = 0
'    lBlkrow2 = 0
'
'End Sub

'Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
'
'    If Row > 0 Then
'        Set Active_Spread = Me.ss1
'        PopupMenu MDIMain.PopUp_Spread
'    End If
'
'End Sub
Private Sub Sp_Zero_Clear()
    Dim i As Integer
    Dim j As Integer
    With ss1
     For i = 1 To .MaxRows
         .Row = i
         For j = 2 To .MaxCols
             .Col = j
             If .CellType = CellTypeNumber And _
             Val(.Text) = 0 Then
                .Text = ""
             End If
         Next j
     Next i
    End With
'    With ss2
'         For i = 1 To .MaxRows
'             .Row = i
'             For j = 2 To .MaxCols
'                 .Col = j
'                 If .CellType = CellTypeNumber And _
'                 Val(.Text) = 0 Then
'                    .Text = ""
'                 End If
'             Next j
'         Next i
'        End With
End Sub

