VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "spr32x60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AQC0030C 
   Caption         =   "材质试验实绩确认 - AQC0030C"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_PIC 
      Caption         =   "复样等待确认"
      Height          =   345
      Left            =   8790
      TabIndex        =   8
      Top             =   120
      Width           =   1275
   End
   Begin InDate.UDate dtp_date_t 
      Height          =   315
      Left            =   3240
      TabIndex        =   5
      Top             =   510
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   556
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
   End
   Begin InDate.UDate dtp_date_f 
      Height          =   315
      Left            =   1620
      TabIndex        =   4
      Top             =   510
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   556
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
   End
   Begin VB.CommandButton cmd_AllCheck 
      Caption         =   "全部确认"
      Height          =   345
      Left            =   7185
      TabIndex        =   3
      Top             =   120
      Width           =   1275
   End
   Begin FPSpread.vaSpread SS1 
      Height          =   7935
      Left            =   195
      TabIndex        =   2
      Top             =   885
      Width           =   6810
      _Version        =   393216
      _ExtentX        =   12012
      _ExtentY        =   13996
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
      MaxCols         =   5
      MaxRows         =   1
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQC0030C.frx":0000
   End
   Begin VB.TextBox txt_SMP_CUT_LOC 
      Height          =   300
      Left            =   5190
      MaxLength       =   1
      TabIndex        =   1
      Top             =   120
      Width           =   1485
   End
   Begin VB.TextBox txt_SMP_NO 
      Height          =   300
      Left            =   1620
      MaxLength       =   14
      TabIndex        =   0
      Top             =   120
      Width           =   1965
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   0
      Left            =   240
      Top             =   120
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   529
      Caption         =   "试样编号"
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
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   1
      Left            =   3810
      Top             =   120
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   529
      Caption         =   "取样位置"
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
      ForeColor       =   -2147483646
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   2
      Left            =   240
      Top             =   510
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   529
      Caption         =   "生产日期"
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
   Begin FPSpread.vaSpread SS2 
      Height          =   7935
      Left            =   7080
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   900
      Width           =   2880
      _Version        =   393216
      _ExtentX        =   5080
      _ExtentY        =   13996
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
      ScrollBars      =   2
      SpreadDesigner  =   "AQC0030C.frx":0506
   End
   Begin FPSpread.vaSpread SS3 
      Height          =   7935
      Left            =   10035
      TabIndex        =   7
      Top             =   900
      Width           =   4890
      _Version        =   393216
      _ExtentX        =   8625
      _ExtentY        =   13996
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   3
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQC0030C.frx":0815
   End
End
Attribute VB_Name = "AQC0030C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       质量管理
'-- Sub_System Name   判定管理
'-- Program Name      材质试验实绩输入
'-- Program ID        AQC0030C
'-- Document No       Q-00-0010(Specification)
'-- Designer          HAN.Y.S
'-- Coder             ZENG.W
'-- Date              2005.10. 25
'-- Description       材质试验实绩输入
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

Dim pColumn12 As New Collection      'Spread Primary Key Collection
Dim nColumn12 As New Collection      'Spread necessary Column Collection
Dim mColumn12 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn12 As New Collection      'Spread Insert Column Collection
Dim aColumn12 As New Collection      'Master -> Spread Column Collection
Dim lColumn12 As New Collection      'Spread Lock Column Collection

Dim pColumn13 As New Collection      'Spread Primary Key Collection
Dim nColumn13 As New Collection      'Spread necessary Column Collection
Dim mColumn13 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn13 As New Collection      'Spread Insert Column Collection
Dim aColumn13 As New Collection      'Master -> Spread Column Collection
Dim lColumn13 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection
Dim Sc3 As New Collection

Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim arrChem(3, 35) As String
Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
      Call Gp_Ms_Collection(txt_SMP_NO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(txt_SMP_CUT_LOC, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(dtp_date_f, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(dtp_date_t, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     
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
     Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AQC0032C.P_REFER", Key:="P-R"
    sc1.Add Item:="AQC0032C.P_MODIFY1", Key:="P-M"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
'     Call SS1.AddCellSpan(5, 0, 1, 2)

      'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", " ", " ", "l", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
     Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", "l", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
     
     'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AQC0032C.P_SREFER_1", Key:="P-R"
    sc2.Add Item:=pColumn12, Key:="pColumn"
    sc2.Add Item:=nColumn12, Key:="nColumn"
    sc2.Add Item:=aColumn12, Key:="aColumn"
    sc2.Add Item:=mColumn12, Key:="mColumn"
    sc2.Add Item:=iColumn12, Key:="iColumn"
    sc2.Add Item:=lColumn12, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"
    
      'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss3, 1, " ", " ", " ", " ", " ", "l", pColumn13, nColumn13, mColumn13, iColumn13, aColumn13, lColumn13)
     Call Gp_Sp_Collection(ss3, 2, " ", " ", " ", " ", " ", "l", pColumn13, nColumn13, mColumn13, iColumn13, aColumn13, lColumn13)
     Call Gp_Sp_Collection(ss3, 3, " ", " ", " ", " ", " ", "l", pColumn13, nColumn13, mColumn13, iColumn13, aColumn13, lColumn13)
     
     'Spread_Collection
    Sc3.Add Item:=ss3, Key:="Spread"
    Sc3.Add Item:="AQC0032C.P_SREFER_2", Key:="P-R"
    Sc3.Add Item:=pColumn13, Key:="pColumn"
    Sc3.Add Item:=nColumn13, Key:="nColumn"
    Sc3.Add Item:=aColumn13, Key:="aColumn"
    Sc3.Add Item:=mColumn13, Key:="mColumn"
    Sc3.Add Item:=iColumn13, Key:="iColumn"
    Sc3.Add Item:=lColumn13, Key:="lColumn"
    Sc3.Add Item:=1, Key:="First"
    Sc3.Add Item:=ss3.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    Call Gp_Sp_ColHidden(ss1, 5, True)
    Call Gp_Sp_ColHidden(ss3, 3, True)
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, 1, ss1.MaxRows, , &HFFFF&)
        
End Sub

Private Sub cmd_AllCheck_Click()
    Dim i       As Integer
    Dim sAllChk As String
    
    If ss1.MaxRows < 1 Or ss1.Row = 0 Then Exit Sub
    
    If cmd_AllCheck.Caption = "全部确认" Then
        sAllChk = "ALL"
    Else
        sAllChk = ""
    End If
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        
        For i = 1 To ss1.MaxRows
            ss1.Row = i
            If sAllChk = "ALL" Then
                ss1.Col = 1
                ss1.Text = 1
                ss1.Col = 0
                ss1.Text = "Update"
                cmd_AllCheck.Caption = "全部取消"
            Else
                ss1.Col = 1
                ss1.Text = 0
                ss1.Col = 0
                ss1.Text = ""
                cmd_AllCheck.Caption = "全部确认"
            End If
        Next i
              
    End If

End Sub

Private Sub MenuToolSet()
     
    MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Insert
    MDIMain.MenuTool.Buttons(8).Enabled = False    'Row Delete
    MDIMain.MenuTool.Buttons(9).Enabled = False    'Row Cancle
    MDIMain.MenuTool.Buttons(11).Enabled = False   'Copy
    MDIMain.MenuTool.Buttons(12).Enabled = False   'Paste
    MDIMain.MenuTool.Buttons(14).Enabled = False
    
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    Call MenuToolSet

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
    Call MenuToolSet

    Call Gp_Ms_Cls(Mc1("rControl"))

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(ss1)
    Call Gp_Sp_Setting(ss2)
    Call Gp_Sp_Setting(ss3)
    Call Gp_Sp_ReadOnlySet(ss2)
    Call Gp_Sp_ReadOnlySet(ss3)

    Call Gf_Sp_Cls(Proc_Sc("Sc"))

    Call Gp_Sp_ColGet(ss1, "Q-System.INI", Me.Name)
    Call Gp_Sp_ColGet(ss2, "Q-System.INI", Me.Name)
    Call Gp_Sp_ColGet(ss3, "Q-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault

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
    
    Set iColumn12 = Nothing
    Set pColumn12 = Nothing
    Set lColumn12 = Nothing
    Set nColumn12 = Nothing
    Set mColumn12 = Nothing
    Set aColumn12 = Nothing
    
    Set iColumn13 = Nothing
    Set pColumn13 = Nothing
    Set lColumn13 = Nothing
    Set nColumn13 = Nothing
    Set mColumn13 = Nothing
    Set aColumn13 = Nothing

    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set sc2 = Nothing
    Set Sc3 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call Gf_Sp_Cls(sc2)
        Call Gf_Sp_Cls(Sc3)
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    End If

End Sub

Public Sub Form_Ref()
    Dim iRow, iCol  As Integer
    Dim sQuery      As String
    Dim sMesg       As String
    Dim AdoRs       As adodb.Recordset

    On Error GoTo Refer_Err
    
    If dtp_date_f.RawData = "" Then
       'dtp_date_f.RawData = Format(Now, "yyyymm") + "01"
       dtp_date_f.RawData = ""
    End If
    
    If dtp_date_t.RawData = "" Then
       dtp_date_t.RawData = Format(Now, "yyyymmdd")
    End If

    If txt_SMP_NO = "" And txt_SMP_CUT_LOC <> "" Then
       MsgBox "请先输入取样号！", vbCritical, "系统提示信息"
       txt_SMP_CUT_LOC = ""
       Exit Sub
    End If

    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        ss1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuToolSet
    End If
    
    Call Gf_Sp_Cls(sc2)
    Call Gf_Sp_Cls(Sc3)
    
    If ss1.MaxRows < 1 Or ss1.ActiveRow = 0 Then Exit Sub
    
    With ss1
         For iRow = 1 To .MaxRows
            .Row = iRow
            .Col = 5
            If .Text = "Y" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HFFFF&)
            Else
                Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &H80000005)
            End If
                  
         Next iRow
    End With
    
Refer_Err:
    
    Screen.MousePointer = vbDefault

End Sub

Public Sub Form_Pro()
'    Dim iRow, iCol As Integer
    Call DataSave("1")

'    If Gf_Sp_Process(M_CN1, Sc1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'
'    SS1.OperationMode = OperationModeNormal
'    Call MenuToolSet
'
'    If SS1.MaxRows < 1 Or SS1.ActiveRow = 0 Then Exit Sub
'    With SS1
'         For iRow = 1 To .MaxRows
'            .Row = iRow
'            .Col = 5
'            If .Text = "Y" Then
'               Call Gp_Sp_BlockColor(SS1, 2, SS1.MaxCols, iRow, iRow, , &HFFFF&)
'            Else
'                Call Gp_Sp_BlockColor(SS1, 2, SS1.MaxCols, iRow, iRow, , &H80000005)
'            End If
'         Next iRow
'    End With

End Sub

Private Sub cmd_PIC_Click()

'    Dim SMP_P As Variant
'
'    With SS1
'        .Row = .ActiveRow
'        .Col = 0
'        .Text = "Update"
'        .Col = 5
'        .Text = "Y"
'        Call Gp_Sp_BlockColor(SS1, 1, SS1.MaxCols, .Row, .Row, , &HFFFF&)
'    End With
   Call DataSave("2")
       
End Sub

Public Sub DataSave(SaveFL As String)
    Dim iRow, iCol As Integer
    
    sc1.Remove ("P-M")
    If SaveFL = "1" Then
        sc1.Add Item:="AQC0032C.P_MODIFY1", Key:="P-M"
    Else
        sc1.Add Item:="AQC0032C.P_MODIFY2", Key:="P-M"
    End If
    
    If Gf_Sp_Process(M_CN1, sc1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    
    ss1.OperationMode = OperationModeNormal
    Call MenuToolSet
    
    If ss1.MaxRows < 1 Or ss1.ActiveRow = 0 Then Exit Sub
    
    With ss1
         For iRow = 1 To .MaxRows
            .Row = iRow
            .Col = 5
            If .Text = "Y" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HFFFF&)
            Else
                Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &H80000005)
            End If
         Next iRow
    End With

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

Public Sub Spread_Del()
    
    Call Gp_Sp_Del(Proc_Sc("SC"))

End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
            
    Dim sQuery          As String
    Dim sMesg           As String
    Dim AdoRs           As adodb.Recordset
    Dim ArrayRecords    As Variant
    Dim arr             As Variant
    Dim smp_no, smp_loc As Variant
 
 'On Error GoTo Error_Rtn
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)

    If ss1.MaxRows < 1 Or Row = 0 Or Col = 1 Then Exit Sub

    With ss1
        .Col = 2
        .Row = .ActiveRow
        smp_no = .Text
        .Col = 3
        smp_loc = .Text
    End With
    
    ss2.MaxRows = 0
    ss3.MaxRows = 0
    
    ss1.ReDraw = False
    ss2.ReDraw = False
    ss3.ReDraw = False
    
    Set AdoRs = New adodb.Recordset
    sQuery = "{call AQC0032C.P_SREFER_1('" + Trim(smp_no) + "')}"

    AdoRs.Open sQuery, M_CN1, adOpenKeyset

    If AdoRs.BOF And AdoRs.EOF Then
        Set AdoRs = Nothing
        'GoTo Error_Rtn
    End If

    ArrayRecords = AdoRs.GetRows
    AdoRs.Close
    
    Call subSpreadView2(ArrayRecords)
    Erase ArrayRecords
    Call Gp_Sp_EvenRowBackcolor(ss2)
        
    sQuery = "{call AQC0032C.P_SREFER_2('" + Trim(smp_no) + "','" + Trim(smp_loc) + "')}"
                    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
    If AdoRs.BOF And AdoRs.EOF Then
        Set AdoRs = Nothing
        GoTo Error_Rtn
    End If
    
    ArrayRecords = AdoRs.GetRows
    AdoRs.Close
    Call subSpreadView1(ArrayRecords)
    Erase ArrayRecords
    
    sQuery = "{call AQC0032C.P_SREFER_3('" + Trim(smp_no) + "')}"
                    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
    If AdoRs.BOF And AdoRs.EOF Then
        Set AdoRs = Nothing
        GoTo Error_Rtn
    End If
    
    ArrayRecords = AdoRs.GetRows
    AdoRs.Close
    Call subSpreadView3(ArrayRecords)
    Erase ArrayRecords

    Call Gp_Sp_EvenRowBackcolor(ss3)
    
    ss1.ReDraw = True
    ss2.ReDraw = True
    ss3.ReDraw = True
    
    Exit Sub
    
Error_Rtn:
    
    Set AdoRs = Nothing
    Screen.MousePointer = vbDefault
    ss1.ReDraw = True
    ss2.ReDraw = True
    ss3.ReDraw = True

End Sub

Private Sub InputEditCheck()

    If ss1.ActiveCol <> 1 Then
        pControl(1).SetFocus
    End If
    
End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)
    Call InputEditCheck
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    Call InputEditCheck
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

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
Private Sub subSpreadView1(ByVal strArr As Variant)

    Dim i           As Integer
    Dim iRow        As Integer
    Dim sMatr(96)   As String
    
    If UBound(strArr, 2) < 0 Then Exit Sub
        
    sMatr(0) = "屈服点实绩                           "
    sMatr(1) = "拉伸规定总伸长应力实绩               "
    sMatr(2) = "抗拉强度实绩                         "
    sMatr(3) = "屈强比实绩                           "
    sMatr(4) = "断后伸长率实绩                       "
    sMatr(5) = "断面收缩率实绩                       "
    sMatr(6) = "冷弯试验实绩                         "
    sMatr(7) = "冲击试验实绩 1                       "
    sMatr(8) = "冲击试验实绩 2                       "
    sMatr(9) = "冲击试验实绩 3                       "
    sMatr(10) = "冲击试验实绩 4                       "
    sMatr(11) = "冲击试验实绩 5                       "
    sMatr(12) = "冲击试验实绩 6                       "
    sMatr(13) = "冲击试验实绩平均                     "
    sMatr(14) = "冲击剪切面积实绩 1                   "
    sMatr(15) = "冲击剪切面积实绩 2                   "
    sMatr(16) = "冲击剪切面积实绩 3                   "
    sMatr(17) = "冲击剪切面积实绩 4                   "
    sMatr(18) = "冲击剪切面积实绩 5                   "
    sMatr(19) = "冲击剪切面积实绩 6                   "
    sMatr(20) = "冲击剪切面积实绩平均                 "
    sMatr(21) = "时效冲击功实绩1                      "
    sMatr(22) = "时效冲击功实绩2                      "
    sMatr(23) = "时效冲击功实绩3                      "
    sMatr(24) = "时效冲击功实绩4                      "
    sMatr(25) = "时效冲击功实绩5                      "
    sMatr(26) = "时效冲击功实绩6                      "
    sMatr(27) = "时效冲击实绩平均                     "
    sMatr(28) = "时效冲击纤维断面率实绩               "
    sMatr(29) = "重力撕裂实绩1                        "
    sMatr(30) = "重力撕裂实绩2                        "
    sMatr(31) = "重力撕裂实绩3                        "
    sMatr(32) = "硬度实绩                             "
    sMatr(33) = "拉伸规定非比例伸长应力实绩           "
    sMatr(34) = "拉伸规定残余伸长应力实绩实绩         "
    sMatr(35) = "高温拉伸屈服强度实绩                 "
    sMatr(36) = "高温拉伸抗拉强度实绩                 "
    sMatr(37) = "高温拉伸断面收缩率实绩               "
    sMatr(38) = "高温拉伸断后伸长率实绩               "
    sMatr(39) = "高温拉伸规定非比例伸长应力实绩       "
    sMatr(40) = "高温拉伸规定残余伸长应力实绩         "
    sMatr(41) = "焊接硬度实绩                         "
    sMatr(42) = "焊缝弯曲实绩                         "
    sMatr(43) = "反复弯曲实绩                         "
    sMatr(44) = "锻平试验实绩                         "
    sMatr(45) = "抗氢裂能力CSR实绩                    "
    sMatr(46) = "抗氢裂能力CLR实绩                    "
    sMatr(47) = "抗氢裂能力CWR实绩                    "
    sMatr(48) = "硫化物腐蚀裂纹实绩                   "
    sMatr(49) = "追加冲击试验实绩 1                   "
    sMatr(50) = "追加冲击试验实绩 2                   "
    sMatr(51) = "追加冲击试验实绩 3                   "
    sMatr(52) = "追加冲击试验实绩 4                   "
    sMatr(53) = "追加冲击试验实绩 5                   "
    sMatr(54) = "追加冲击试验实绩 6                   "
    sMatr(55) = "追加冲击试验实绩平均                 "
    sMatr(56) = "追加冲击剪切面积实绩 1               "
    sMatr(57) = "追加冲击剪切面积实绩 2               "
    sMatr(58) = "追加冲击剪切面积实绩 3               "
    sMatr(59) = "追加冲击剪切面积实绩 4               "
    sMatr(60) = "追加冲击剪切面积实绩 5               "
    sMatr(61) = "追加冲击剪切面积实绩 6               "
    sMatr(62) = "追加冲击剪切面积实绩平均             "
    sMatr(63) = "追加时效冲击功实绩1                  "
    sMatr(64) = "追加时效冲击功实绩2                  "
    sMatr(65) = "追加时效冲击功实绩3                  "
    sMatr(66) = "追加时效冲击功实绩4                  "
    sMatr(67) = "追加时效冲击功实绩5                  "
    sMatr(68) = "追加时效冲击功实绩6                  "
    sMatr(69) = "追加时效冲击实绩平均                 "
    sMatr(70) = "追加时效冲击纤维断面率实绩           "
    sMatr(71) = "晶粒度实绩                           "
    sMatr(72) = "脱碳层实绩                           "
    sMatr(73) = "硫印实绩                             "
    sMatr(74) = "断口检验实绩1                        "
    sMatr(75) = "断口检验实绩2                        "
    sMatr(76) = "断口检验实绩3                        "
    sMatr(77) = "断口检验实绩4                        "
    sMatr(78) = "断口检验实绩5                        "
    sMatr(79) = "酸浸检验实绩1                        "
    sMatr(80) = "酸浸检验实绩2                        "
    sMatr(81) = "酸浸检验实绩3                        "
    sMatr(82) = "酸浸检验实绩4                        "
    sMatr(83) = "酸浸检验实绩5                        "
    sMatr(84) = "带状组织实绩                         "
    sMatr(85) = "淬透性试验实绩1                      "
    sMatr(86) = "淬透性试验实绩2                      "
    sMatr(87) = "淬透性试验实绩3                      "
    sMatr(88) = "非金属夹杂物(粗)实绩1                "
    sMatr(89) = "非金属夹杂物(粗)实绩2                "
    sMatr(90) = "非金属夹杂物(粗)实绩3                "
    sMatr(91) = "非金属夹杂物(粗)实绩4                "
    sMatr(92) = "非金属夹杂物(细)实绩1                "
    sMatr(93) = "非金属夹杂物(细)实绩2                "
    sMatr(94) = "非金属夹杂物(细)实绩3                "
    sMatr(95) = "非金属夹杂物(细)实绩4                "
  
    With ss3
        .MaxRows = 96
    
        For i = 1 To 96
            .Row = i
            .Col = 1: .Text = sMatr(i - 1)
        Next i
                
        For i = 1 To UBound(strArr, 1) + 1
        
            .Row = i: .Col = 2
            .Text = NullCheck(strArr(i - 1, 0), "")
            
        Next i
    End With

End Sub

Private Sub subSpreadView3(ByVal strArr As Variant)

    Dim i       As Integer
    Dim iRow    As Integer
    
    If UBound(strArr, 2) < 0 Then Exit Sub
      
    With ss3
        .MaxRows = 96
        
        For i = 1 To UBound(strArr, 1) + 1
        
            .Row = i: .Col = 3
            
            .Text = NullCheck(strArr(i - 1, 0), "")
            
        Next i
    End With
     Call subSpreadCheck1
End Sub

Private Sub subSpreadView2(ByVal strArr As Variant)

    Dim i       As Integer
    Dim iRow    As Integer

    With ss2

        .MaxRows = 0
        .MaxRows = 33

'        For i = 1 To 17
'            .Row = i
'            .Col = 1: .Text = sChem(i - 1)
'        Next i

        For i = 1 To UBound(strArr, 2) + 1
            .Row = i: .Col = 1

            .Text = NullCheck(strArr(0, i - 1), "")
            
            .Row = i: .Col = 2

            .Text = NullCheck(strArr(1, i - 1), "")
        Next i

    End With
    Call subSpreadCheck2
    
End Sub

Private Sub subSpreadCheck2()
    
    Dim i As Long
    Dim J As Long
    
    With ss2
        
        For i = 1 To 33
                                    
            If Gf_Get_Cell_Value(ss2, i, 2) = "" Or Gf_Get_Cell_Value(ss2, i, 2) = 0 Then
                .Row = i
                .RowHidden = True
            Else
                .RowHidden = False
                J = J + 1
                .Col = 0: .Text = J
            End If
        Next i
                
    End With
    
End Sub

Private Sub subSpreadCheck1()
    
    Dim i As Long
    Dim J As Long
    
    With ss3
       
       For i = 1 To 96
                                   
           If (Gf_Get_Cell_Value(ss3, i, 2) = "" Or Gf_Get_Cell_Value(ss3, i, 2) = 0 _
               Or IsNull(Gf_Get_Cell_Value(ss3, i, 2))) And Gf_Get_Cell_Value(ss3, i, 3) = "" Then
               .Row = i
               .RowHidden = True
           Else
               .RowHidden = False
               J = J + 1
               .Col = 0: .Text = J
           End If
       Next i
               
    End With
End Sub

Private Sub txt_SMP_CUT_LOC_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
'        If txt_SMP_NO = "" Then
'           MsgBox "请先输入取样号！", vbCritical, "系统提示信息"
'           txt_SMP_CUT_LOC = ""
'           Exit Sub
'        End If

        DD.sWitch = "MS"
        DD.sKey = "Q0042"
        DD.rControl.Add Item:=txt_SMP_CUT_LOC

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

End Sub

Private Sub ss1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

    If Row < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        With ss1
            .Row = Row
            .Col = 5
            If .BackColor = &HFFFF& Then
                .Col = 1
                If .Text = "1" Then
                    .Col = 5:   .Text = ""
                    .Col = 0:   .Text = "Update"
                Else
                    .Col = 5:   .Text = "Y"
                    .Col = 0:   .Text = ""
                End If
            Else
                .Col = 1
                If .Text = "1" Then
                    .Col = 5:   .Text = "Y"
                    .Col = 0:   .Text = "Update"
                Else
                    .Col = 5:   .Text = ""
                    .Col = 0:   .Text = ""
                End If
            End If
        End With
    End If
    
End Sub
