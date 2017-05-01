VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form ABX1070C 
   Caption         =   "厚度宽度组_ABX1070C"
   ClientHeight    =   7620
   ClientLeft      =   1380
   ClientTop       =   4020
   ClientWidth     =   13365
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7620
   ScaleWidth      =   13365
   WindowState     =   2  'Maximized
   Begin VB.OptionButton Opt1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "厚度组"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   150
      Width           =   1215
   End
   Begin VB.OptionButton Opt2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "宽度组"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1860
      TabIndex        =   1
      Top             =   150
      Width           =   1260
   End
   Begin VB.TextBox txt_prod_cd 
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
      Left            =   4710
      MaxLength       =   2
      TabIndex        =   2
      Tag             =   "产品代码"
      Top             =   135
      Width           =   705
   End
   Begin VB.TextBox txt_prod_name 
      Enabled         =   0   'False
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
      Left            =   5415
      MaxLength       =   40
      TabIndex        =   3
      Tag             =   "Prod_Cd_Name"
      Top             =   135
      Width           =   2670
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   3345
      Top             =   135
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "产品"
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
   Begin FPSpread.vaSpread ss2 
      Height          =   8640
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   510
      Width           =   15045
      _Version        =   393216
      _ExtentX        =   26538
      _ExtentY        =   15240
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
      MaxCols         =   8
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "ABX1070C.frx":0000
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8640
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   510
      Width           =   15045
      _Version        =   393216
      _ExtentX        =   26538
      _ExtentY        =   15240
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
      MaxCols         =   8
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "ABX1070C.frx":1AA8
   End
End
Attribute VB_Name = "ABX1070C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       System Management
'-- Sub_System Name   Code Management
'-- Program Name      Common Code
'-- Program ID        AZA1010C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
'-- Date              2003.5.19
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

Dim pControl2 As New Collection     'Master Primary Key Collection
Dim nControl2 As New Collection     'Master Necessary Collection
Dim mControl2 As New Collection     'Master Maxlength check Collection
Dim iControl2 As New Collection     'Master Insert Collection
Dim rControl2 As New Collection     'Master Refer Collection
Dim cControl2 As New Collection     'Master Copy Collection
Dim aControl2 As New Collection     'Master -> Spread Collection
Dim lControl2 As New Collection     'Master Lock Collection

Dim pColumn1 As New Collection      'Spread Primary Key Collection
Dim nColumn1 As New Collection      'Spread necessary Column Collection
Dim mColumn1 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn1 As New Collection      'Spread Insert Column Collection
Dim aColumn1 As New Collection      'Master -> Spread Column Collection
Dim lColumn1 As New Collection      'Spread Lock Column Collection

Dim pColumn2 As New Collection      'Spread Primary Key Collection
Dim nColumn2 As New Collection      'Spread necessary Column Collection
Dim mColumn2 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn2 As New Collection      'Spread Insert Column Collection
Dim aColumn2 As New Collection      'Master -> Spread Column Collection
Dim lColumn2 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()

   ' Dim sQuery As String
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------  BELOW EDIT 1      ---------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------------------

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
           Call Gp_Ms_Collection(txt_prod_cd, "P", "n", " ", " ", "r", "a", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_prod_name, " ", " ", " ", " ", " ", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------  EDIT 1 End      ---------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
  
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    'MASTER Collection
    Mc2.Add Item:=pControl2, Key:="pControl"
    Mc2.Add Item:=nControl2, Key:="nControl"
    Mc2.Add Item:=mControl2, Key:="mControl"
    Mc2.Add Item:=iControl2, Key:="iControl"
    Mc2.Add Item:=rControl2, Key:="rControl"
    Mc2.Add Item:=cControl2, Key:="cControl"
    Mc2.Add Item:=aControl2, Key:="aControl"
    Mc2.Add Item:=lControl2, Key:="lControl"

'---------------------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------  BELOW EDIT 2      ---------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------------------

    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", "i", "a", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
     Call Gp_Sp_Collection(ss2, 1, "p", " ", " ", "i", "a", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 2, "p", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="ABX1070C.P_MODIFY1", Key:="P-M"
    sc1.Add Item:="ABX1070C.P_REFER1", Key:="P-R"
    sc1.Add Item:="ABX1070C.P_ONEROW1", Key:="P-O"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="ABX1070C.P_MODIFY2", Key:="P-M"
    sc2.Add Item:="ABX1070C.P_ONEROW2", Key:="P-O"
    sc2.Add Item:="ABX1070C.P_REFER2", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    ss1.Col = 7
    ss1.ColHidden = True
    ss2.Col = 7
    ss2.ColHidden = True
'    Sc1.Item("Spread").Col = 0
'    Sc1.Item("Spread").Row = 0
'    Sc1.Item("Spread").Text = "⒗"

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
   ' Call Gp_Ms_Cls(Mc2("rControl"))
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(sc1.Item("Spread"))
    Call Gp_Sp_Setting(sc2.Item("Spread"))
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "B-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "B-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault
    
    Opt1.Value = True
    ss1.Visible = True
    ss2.Visible = False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(sc1.Item("Spread"), "B-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "B-System.INI", Me.Name)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set pControl2 = Nothing
    Set nControl2 = Nothing
    Set iControl2 = Nothing
    Set rControl2 = Nothing
    Set cControl2 = Nothing
    Set aControl2 = Nothing
    Set lControl2 = Nothing
    Set mControl2 = Nothing
    
    Set iColumn1 = Nothing
    Set pColumn1 = Nothing
    Set lColumn1 = Nothing
    Set nColumn1 = Nothing
    Set mColumn1 = Nothing
    Set aColumn1 = Nothing
    
    Set iColumn2 = Nothing
    Set pColumn2 = Nothing
    Set lColumn2 = Nothing
    Set nColumn2 = Nothing
    Set mColumn2 = Nothing
    Set aColumn2 = Nothing
    
    Set Mc1 = Nothing
    Set Mc2 = Nothing
    Set sc1 = Nothing
    Set sc2 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, sc1)
    Call Gp_Sp_Cancel(M_CN1, sc2)
      
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(sc2) Then
        If Gf_Sp_Cls(sc1) Then
            Call Gp_Ms_Cls(Mc1("rControl"))
       '    Call Gp_Ms_Cls(Mc2("rControl"))
            Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
            Call Gp_Ms_ControlLock(Mc1("lControl"), False)
            rControl(1).SetFocus
            Opt1.Value = True
            Opt1.Enabled = True
            Opt2.Enabled = True
            ss1.Visible = True
            ss2.Visible = False
            txt_prod_name = ""
        End If
    End If
    
End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    Dim SMESG As String
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
            
    If Opt1.Value = True Then
        Call Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Gf_Sp_Cls(sc2)
        Opt2.Enabled = False
        Exit Sub
    Else
        Call Gf_Sp_Refer(M_CN1, sc2, Mc1, Mc1("nControl"), Mc1("mControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Gf_Sp_Cls(sc1)
        Opt1.Enabled = False
        Exit Sub
    End If
            
    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Pro()
    
    If Opt1.Value = True Then
    
        If ss_check(sc1.Item("Spread")) Then
           Call Gf_Sp_Process(M_CN1, sc1, Mc1)
           Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        End If
        
    End If
    
    If Opt2.Value = True Then
        
        If ss_check(sc2.Item("Spread")) Then
           Call Gf_Sp_Process(M_CN1, sc2, Mc1)
           Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        End If
           
    End If

End Sub

Public Sub Form_Ins()
    
    If Opt1.Value = True Then
        If ss1.ActiveRow <> ss1.MaxRows Then Exit Sub
        
        Call Gp_Sp_Ins(sc1)
        
        Call Gp_Sp_InAuthority(sc1, 7)
        ss1.Col = 1
        ss1.Text = txt_prod_cd.Text
        ss1.Col = 4
        
    Else
        If ss2.ActiveRow <> ss2.MaxRows Then Exit Sub
        Call Gp_Sp_Ins(sc2)
        Call Gp_Sp_InAuthority(sc2, 7)
        ss2.Col = 1
        ss2.Text = txt_prod_cd.Text
        ss2.Col = 4
    
       
    End If

End Sub

Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(sc1)
    Call Gp_Sp_Copy(sc2)
    
End Sub

Public Sub Spread_Pst()

    Call Gp_Sp_Paste(sc1)
    Call Gp_Sp_Paste(sc2)
    
End Sub

Public Sub Spread_ColumnsSort()

'    Spread_ColSort.Show 1
    
End Sub

Public Sub Spread_Forzens_Setting()
    
    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol
    
End Sub

Public Sub Spread_Forzens_Cancel()

    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = 0
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub Spread_Del()
    
    Call Gp_Sp_Del(sc1)
    Call Gp_Sp_Del(sc2)
End Sub

Private Sub Opt1_Click()

    Opt2.Value = False
'    Opt2.Enabled = False
    Opt1.Value = True
    ss2.Visible = False
    ss1.Visible = True
 
    txt_prod_cd = ""
    txt_prod_name = ""
    
End Sub

Private Sub Opt2_Click()

    Opt1.Value = False
'    Opt1.Enabled = False
    Opt2.Value = True
    ss1.Visible = False
    ss2.Visible = True
 
    txt_prod_cd = ""
    txt_prod_name = ""

End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
 '   Call Gp_Sp_Sort(sc1.Item("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 7)
    End If
    
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 7)
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub

Private Sub ss1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

'   Max Thick check

    Dim iCol As Integer
    Dim iRow As Integer
    Dim DCURR As Double
    Dim dThick As Double
    If Col = 1 Or Row <= 2 Then Exit Sub
            
    With ss1
    
        If .CellTag = "False" Then Exit Sub
        
        .Col = 4
        .Row = Row
        If .Value = "" Then
           dThick = 0
        Else
           dThick = .Value
        End If
        
        .Col = 4
        .Row = Row - 1
        If .Value = "" Then
            DCURR = 0
        Else
            DCURR = .Value
        End If
         
        
        If DCURR >= dThick Or DCURR = 0 Then
           .Col = Col
           .Row = Row
           .CellTag = "False"
           Call Gp_MsgBoxDisplay("数据有错,检查后重新输入...")
           .Col = Col
           .Row = Row
           .CellTag = ""
           .Value = 0
           .TabStop = True
           .SetFocus
           .SetActiveCell Col, Row
           .Action = SS_ACTION_ACTIVE_CELL
           .EditMode = True
           .TabStop = False
            
           Exit Sub
        End If
            
    End With
End Sub

Private Sub ss1_LeaveRow(ByVal Row As Long, ByVal RowWasLast As Boolean, ByVal RowChanged As Boolean, ByVal AllCellsHaveData As Boolean, ByVal NewRow As Long, ByVal NewRowIsLast As Long, Cancel As Boolean)

'   Max Thick check

    Dim iCol As Integer
    Dim iRow As Integer
    Dim DCURR As Double
    Dim dThick As Double
    If Row <= 2 Then Exit Sub
            
    With ss1
    
        If .CellTag = "False" Then Exit Sub
        
        .Col = 4
        .Row = Row
        If .Value = "" Then
           dThick = 0
        Else
           dThick = .Value
        End If
        
        .Col = 4
        .Row = Row - 1
        If .Value = "" Then
            DCURR = 0
        Else
            DCURR = .Value
        End If
         
        
        If DCURR >= dThick Or DCURR = 0 Then
           .Col = 4
           .Row = Row
           .CellTag = "False"
           Call Gp_MsgBoxDisplay("数据有错,检查后重新输入...")
           .Col = 4
           .Row = Row
           .CellTag = ""
           .Value = 0
           .TabStop = True
           .SetFocus
           .SetActiveCell 4, Row
           .Action = SS_ACTION_ACTIVE_CELL
           .EditMode = True
           .TabStop = False
            
           Exit Sub
        End If
            
    End With


End Sub

Private Sub ss2_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

'   Max Width check

    Dim iCol As Integer
    Dim iRow As Integer
    Dim DCURR As Double
    Dim dWidth As Double
    
    If Col = 1 Or Row <= 2 Then Exit Sub
            
    With ss2
    
        If .CellTag = "False" Then Exit Sub
        
        .Col = 4
        .Row = Row
        If .Value = "" Then
           dWidth = 0
        Else
           dWidth = .Value
        End If
        
        .Col = 4
        .Row = Row - 1
        If .Value = "" Then
            DCURR = 0
        Else
            DCURR = .Value
        End If
         
        
        If DCURR >= dWidth Or DCURR = 0 Then
           .Col = Col
           .Row = Row
           .CellTag = "False"
           Call Gp_MsgBoxDisplay("数据有错,检查后重新输入...")
           .Col = Col
           .Row = Row
           .CellTag = ""
           .Value = 0
           .TabStop = True
           .SetFocus
           .SetActiveCell Col, Row
           .Action = SS_ACTION_ACTIVE_CELL
           .EditMode = True
           .TabStop = False
             
           Exit Sub
        End If
            
    End With
   
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


Private Sub ss2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)
    
'    Call Gp_Sp_Sort(sc2.Item("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("Sc")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 7)
    End If
    
End Sub

Private Sub ss2_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 7)
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub


Private Sub ss2_LeaveRow(ByVal Row As Long, ByVal RowWasLast As Boolean, ByVal RowChanged As Boolean, ByVal AllCellsHaveData As Boolean, ByVal NewRow As Long, ByVal NewRowIsLast As Long, Cancel As Boolean)
'   Max Width check

    Dim iCol As Integer
    Dim iRow As Integer
    Dim DCURR As Double
    Dim dWidth As Double
    
    If Row <= 2 Then Exit Sub
            
    With ss2
    
        If .CellTag = "False" Then Exit Sub
        
        .Col = 4
        .Row = Row
        If .Value = "" Then
           dWidth = 0
        Else
           dWidth = .Value
        End If
        
        .Col = 4
        .Row = Row - 1
        If .Value = "" Then
            DCURR = 0
        Else
            DCURR = .Value
        End If
         
        
        If DCURR >= dWidth Or DCURR = 0 Then
           .Col = 4
           .Row = Row
           .CellTag = "False"
           Call Gp_MsgBoxDisplay("数据有错,检查后重新输入...")
           .Col = 4
           .Row = Row
           .CellTag = ""
           .Value = 0
           .TabStop = True
           .SetFocus
           .SetActiveCell 4, Row
           .Action = SS_ACTION_ACTIVE_CELL
           .EditMode = True
           .TabStop = False
             
           Exit Sub
        End If
            
    End With
   
End Sub

Private Sub ss2_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss2_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss2
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub

Private Sub txt_prod_cd_DblClick()
Call txt_prod_cd_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_prod_cd_KeyPress(KeyAscii As Integer)

     KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub

Private Sub txt_prod_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0005"
        DD.rControl.Add Item:=txt_prod_cd
        DD.rControl.Add Item:=txt_prod_name

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_prod_cd)) = txt_prod_cd.MaxLength Then
        txt_prod_name.Text = Gf_ComnNameFind(M_CN1, "B0005", Trim(txt_prod_cd.Text), 2)
    Else
        txt_prod_name.Text = ""
    End If

End Sub

Private Sub ss1_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sTemp_Code As String

    If ss1.MaxRows < 1 Then Exit Sub
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        Exit Sub
    End If

    Select Case ss1.ActiveCol
    
        Case 1
        
            If KeyCode = vbKeyF4 Then
            
                Set DD.sPname = Me.ss1
                
                DD.sWitch = "SP"
                DD.sKey = "B0005"
                DD.rControl.Add Item:=1
  '              DD.rControl.Add Item:=2
                
                DD.nameType = "2"
                Call Gf_Common_DD(M_CN1, KeyCode)
                
            End If
            
    End Select
    
End Sub

Private Sub ss2_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sTemp_Code As String

    If ss2.MaxRows < 1 Then Exit Sub
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        Exit Sub
    End If

    Select Case ss2.ActiveCol
    
        Case 1
        
            If KeyCode = vbKeyF4 Then
            
                Set DD.sPname = Me.ss2
                
                DD.sWitch = "SP"
                DD.sKey = "B0005"
                DD.rControl.Add Item:=1
         '       DD.rControl.Add Item:=2
                
                DD.nameType = "2"
                Call Gf_Common_DD(M_CN1, KeyCode)
            
            End If
            
    End Select
    
End Sub

Private Function ss_check(sPname As Variant) As Boolean

 On Error GoTo ss_check_Error
 
    Dim iCol As Integer
    Dim iRow As Integer
    Dim DCURR As Double
    Dim dLast As Double
    
    ss_check = True
            
    With sPname
        
        If .MaxRows <= 2 Then Exit Function
        
        If .CellTag = "False" Then Exit Function
        
        .Col = 4
        .Row = .MaxRows
        If .Value = "" Then
           dLast = 0
        Else
           dLast = .Value
        End If
        
        .Col = 4
        .Row = .MaxRows - 1
        If .Value = "" Then
            DCURR = 0
        Else
            DCURR = .Value
        End If
         
        
        If DCURR >= dLast Or DCURR = 0 Then
           .Col = 4
           .Row = .MaxRows
           .CellTag = "False"
           Call Gp_MsgBoxDisplay("数据有错,检查后重新输入...")
           .Col = 4
           .Row = .MaxRows
           .CellTag = ""
           .Value = 0
           .TabStop = True
           .SetFocus
           .SetActiveCell 4, .MaxRows
           .Action = SS_ACTION_ACTIVE_CELL
           .EditMode = True
           .TabStop = False
           
           ss_check = False
               
          
        Else
           ss_check = True
           
        End If
        Exit Function
        
    End With

ss_check_Error:

    ss_check = False
    
End Function
