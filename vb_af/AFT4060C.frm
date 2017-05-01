VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form AFT4060C 
   Caption         =   "板坯判定实绩修改及查询界面_AFT4060C"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   WindowState     =   2  'Maximized
   Begin Threed.SSPanel SSPanel1 
      Height          =   1050
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   20250
      _ExtentX        =   35719
      _ExtentY        =   1852
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
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
         ItemData        =   "AFT4060C.frx":0000
         Left            =   10350
         List            =   "AFT4060C.frx":0002
         TabIndex        =   12
         Tag             =   "机号"
         Top             =   120
         Width           =   735
      End
      Begin VB.TextBox TXT_PD_FL 
         Height          =   270
         Left            =   3240
         TabIndex        =   10
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   11760
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txt_slab_no 
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
         Left            =   6735
         MaxLength       =   10
         TabIndex        =   1
         Top             =   120
         Width           =   1425
      End
      Begin InDate.UDate txt_to_DATE 
         Height          =   315
         Left            =   3045
         TabIndex        =   2
         Tag             =   "终止日期"
         Top             =   120
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
      Begin InDate.UDate txt_from_DATE 
         Height          =   315
         Left            =   1335
         TabIndex        =   3
         Tag             =   "起始日期"
         Top             =   120
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
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Left            =   5520
         Top             =   120
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "板坯号"
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
      Begin Threed.SSOption OPT_NOORD 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   255
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "未判定"
         Value           =   -1
      End
      Begin Threed.SSOption OPT_INSPSCRAP 
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   16711680
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "已判定"
      End
      Begin InDate.ULabel ULabel4 
         Height          =   315
         Left            =   9120
         Top             =   120
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "铸机号"
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
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00E1E4CD&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "生产日期"
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
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "～"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2805
         TabIndex        =   4
         Top             =   120
         Width           =   375
      End
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   3030
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   5345
      _Version        =   196609
      AutoSize        =   1
      PaneTree        =   "AFT4060C.frx":0004
      Begin FPSpread.vaSpread ss1 
         Height          =   1935
         Left            =   30
         TabIndex        =   6
         Top             =   1065
         Width           =   4500
         _Version        =   393216
         _ExtentX        =   7937
         _ExtentY        =   3413
         _StockProps     =   64
         ColHeaderDisplay=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   22
         MaxRows         =   3
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AFT4060C.frx":0056
      End
   End
End
Attribute VB_Name = "AFT4060C"
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
'-- Program Name      板坯库存报表
'-- Program ID        AFT4040C
'-- Designer          wanglei
'-- Coder             wanglei
'-- Date              2014.6.24
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
Public QueryYN As Boolean

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
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection


Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
        
    Dim iCol As Integer
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
    Call Gp_Ms_Collection(txt_from_DATE, "p ", "n ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_to_DATE, "p ", "n ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   
    Call Gp_Ms_Collection(txt_slab_no, "p ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(TXT_PD_FL, "p ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(cbo_prc_line, "p ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  
  
  
   
   
    
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
    
        Call Gp_Sp_Collection(ss1, 1, "p", "n", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
        Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
        Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
        Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
        Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
        Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
        Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
        Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
        Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
        Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
        Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
        Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
        Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
        Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
        Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
        Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
        Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
        Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
        Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
        Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
        Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
        Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
        
   
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AFT4060C.P_SREFER", Key:="P-R"
    sc1.Add Item:="AFT4060C.P_ONEROW", Key:="P-O"
    sc1.Add Item:="AFT4060C.P_MODIFY", Key:="P-M"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxRows, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").Row = 0
    sc1.Item("Spread").Text = "◎"
    

    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)

End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    cbo_prc_line.AddItem ""
    cbo_prc_line.AddItem "1"
    cbo_prc_line.AddItem "2"
    cbo_prc_line.AddItem "3"
     
    Call Form_Define
  
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "F-System.INI", Me.Name)
    txt_from_DATE.RawData = Format(Now, "YYYYMMDD")
    txt_to_DATE.RawData = Format(Now, "YYYYMMDD")
    Screen.MousePointer = vbDefault
    TXT_PD_FL = "4"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "F-System.INI", Me.Name)
    
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
    Set sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        txt_from_DATE.RawData = Format(Now, "YYYYMMDD")
        txt_to_DATE.RawData = Format(Now, "YYYYMMDD")
'        txt_OCCUR_DATE.SetFocus
    End If

End Sub

Public Sub Form_Ref()
    Dim i, j As Integer

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1) Then
        ss1.SetFocus
        ss1.OperationMode = OperationModeNormal
        Call Gp_Sp_EvenRowBackcolor(Proc_Sc("Sc")("Spread"))
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
    
    For i = 15 To 20
        For j = 1 To ss1.MaxRows
            ss1.Col = i
            ss1.Row = j
            If Len(Trim(ss1.Text)) = 4 Then
            ss1.Text = Gf_ComnNameFind(M_CN1, "F0045", ss1.Text, 2)
            End If
        Next j
    Next i
            
End Sub

Private Sub OPT_INSPSCRAP_Click(VALUE As Integer)
    
    
    If OPT_INSPSCRAP.VALUE = True Then
        OPT_INSPSCRAP.ForeColor = &HFF&
        OPT_NOORD.ForeColor = &H808080
        Label2.Caption = "判定日期"
        TXT_PD_FL = "1"
    Else
        OPT_INSPSCRAP.ForeColor = &H808080
        TXT_PD_FL = "4"
    End If
End Sub

Private Sub OPT_NOORD_Click(VALUE As Integer)
    Dim iRow As Integer
    Dim sTemp As String
    
    If OPT_NOORD.VALUE = True Then
        OPT_NOORD.ForeColor = &HFF&
        OPT_INSPSCRAP.ForeColor = &H808080
        Label2.Caption = "生产日期"
        TXT_PD_FL = "4"
    Else
        OPT_NOORD.ForeColor = &H808080
        TXT_PD_FL = "1"
    End If
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
If Col >= 2 Then
    ss1.Col = 0
    ss1.Row = Row
    Select Case Trim(ss1.Text)
           Case "Input", "Update", "Delete"
           Case Else
                ss1.Text = "Update"
    End Select
      
    With ss1
        .Row = .ActiveRow
        .Col = 22
        .VALUE = Format(Now, "YYYYMMDDHHMMSS")
        .Col = 21
        .Text = sUserID
    End With
    
End If


End Sub
'Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
'
''    If Gf_Sc_Authority(sAuthority, "U") Then Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
'
'    If Gf_Sc_Authority(sAuthority, "U") Then
'        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
'    End If
'
'End Sub



Public Sub Form_Pro()
    

    If Gf_Mc_Authority(sAuthority, Mc1) Then
        If Gf_Sp_Process(M_CN1, Proc_Sc("Sc"), Mc1) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
            
        End If
    End If

    Call Form_Ref

    
End Sub

Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    
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

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
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
'Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
'
'    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
'
'    lBlkcol1 = 0
'    lBlkcol2 = 0
'    lBlkrow1 = 0
'    lBlkrow2 = 0
'
'End Sub
Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

With ss1

    If Col = 2 Then
       
       .Col = .ActiveCol
       .Row = .ActiveRow
                
        If .VALUE = 0 Then
           .Text = "0"
         
           
        ElseIf .VALUE = 1 Then
           .Text = "1"
           
        End If
    
    End If
    
End With

End Sub
Private Sub SS1_KeyUp(KeyCode As Integer, Shift As Integer)

    If ss1.MaxRows <= 0 Then Exit Sub
    If ss1.ActiveRow <= 0 Then Exit Sub
    If ss1.ActiveCol >= 15 And ss1.ActiveCol <= 20 Then
      If KeyCode = vbKeyF4 Then
        With ss1
             .Row = .ActiveRow

             .Col = .ActiveCol
             ss1.SetFocus
             Text1.Text = .Text
        End With
        Call txt_pd_cd_KeyUp(vbKeyF4, 0)
       End If

       With ss1
             .Row = .ActiveRow

             .Col = .ActiveCol
            
             Text1.Text = .Text
            
       End With
       
       If Len(Text1.Text) = 4 Then
         If IsNumeric(Text1.Text) Then
             With ss1
             .Row = .ActiveRow
             .Col = .ActiveCol
             .Text = Gf_ComnNameFind(M_CN1, "F0045", Text1.Text, 2)
             End With
         End If
        End If

    End If
     
   
End Sub
'Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)
'
'
'    If Row <= 0 Then Exit Sub
'    If ss1.ActiveCol >= 15 And ss1.ActiveCol <= 20 Then
'
'        With ss1
'             .Row = .ActiveRow
'
'             .Col = .ActiveCol
'             ss1.SetFocus
'             Text1.Text = .Text
'        End With
'        Call txt_pd_cd_KeyUp(vbKeyF4, 0)
'    End If
'End Sub
Private Sub txt_pd_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sMesg  As String
    Dim sQuery As String
    
    If KeyCode = vbKeyF4 Then
        
        DD.sWitch = "MS"
        DD.sKey = "F0045"
        DD.rControl.Add Item:=Text1
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
    End If
    
'
    With ss1
         .Row = .ActiveRow
'
         .Col = .ActiveCol
         .Text = Text1.Text
         
         .Text = Gf_ComnNameFind(M_CN1, "F0045", Text1.Text, 2)
         .Col = 0
         .Text = "Update"
         .Col = 22
         .VALUE = Format(Now, "YYYYMMDDHHMMSS")
         .Col = 21
         .Text = sUserID
         
    End With
   
        
End Sub
Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

'Private Sub txt_OCCUR_DATE_DblClick()
'    txt_OCCUR_DATE.RawData = Format(Now, "YYYYMMDD")
'End Sub



