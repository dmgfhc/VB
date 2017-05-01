VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form AKP1014C 
   Caption         =   "板坯质量查询_AKP1014C"
   ClientHeight    =   9225
   ClientLeft      =   510
   ClientTop       =   2325
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15225
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8685
      Left            =   60
      TabIndex        =   7
      Top             =   510
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   15319
      _Version        =   196609
      SplitterBarWidth=   4
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "AKP1014C.frx":0000
      Begin FPSpread.vaSpread ss1 
         Height          =   6135
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   15105
         _Version        =   393216
         _ExtentX        =   26644
         _ExtentY        =   10821
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
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
         MaxCols         =   19
         MaxRows         =   1
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKP1014C.frx":0052
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   2490
         Left            =   0
         TabIndex        =   9
         Top             =   6195
         Width           =   15105
         _Version        =   393216
         _ExtentX        =   26644
         _ExtentY        =   4392
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
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
         MaxCols         =   17
         MaxRows         =   1
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKP1014C.frx":0B3E
      End
   End
   Begin VB.TextBox TXT_BOF_CC 
      Height          =   330
      Left            =   7860
      MaxLength       =   1
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "1"
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin InDate.ULabel ubl_Date 
      Height          =   315
      Left            =   210
      Top             =   135
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Caption         =   "生产日期"
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
   Begin InDate.UDate txt_to_DATE 
      Height          =   315
      Left            =   3045
      TabIndex        =   1
      Tag             =   "终止日期"
      Top             =   135
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
   Begin InDate.UDate txt_from_DATE 
      Height          =   315
      Left            =   1275
      TabIndex        =   0
      Tag             =   "起始日期"
      Top             =   135
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
   Begin Threed.SSCommand cmd_Edit 
      Height          =   420
      Left            =   13740
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   60
      Width           =   1365
      _ExtentX        =   2408
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
      Caption         =   "班组查询"
      BevelWidth      =   3
   End
   Begin Threed.SSOption opt_bof_date 
      Height          =   285
      Left            =   4770
      TabIndex        =   2
      Top             =   150
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "冶炼时间"
      Value           =   -1
   End
   Begin Threed.SSOption opt_cc_date 
      Height          =   285
      Left            =   6300
      TabIndex        =   3
      Top             =   150
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   8421504
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
      Caption         =   "连铸时间"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "～"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2805
      TabIndex        =   5
      Top             =   165
      Width           =   255
   End
End
Attribute VB_Name = "AKP1014C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       Production DayReport Final Steel Grade
'-- Sub_System Name
'-- Program Name
'-- Program ID        BFC2014C
'-- Document No       Q-00-0010(Specification)
'-- Designer
'-- Coder
'-- Date              2005.06.14
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

Dim pColumn2 As New Collection      'Spread Primary Key Collection
Dim nColumn2 As New Collection      'Spread necessary Column Collection
Dim mColumn2 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn2 As New Collection      'Spread Insert Column Collection
Dim aColumn2 As New Collection      'Master -> Spread Column Collection
Dim lColumn2 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim ls_PChangeName                  'To Record P control Name

Private Sub Form_Define()
   Dim i As Integer
   
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
  Call Gp_Ms_Collection(txt_from_DATE, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_to_DATE, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_BOF_CC, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     
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
   
   For i = 1 To ss1.MaxCols
      Call Gp_Sp_ColColor(ss1, i, , &H8000000B)
   Next i
   
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AKP1014C.P_SREFER1", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    Call Gp_Sp_ColHidden(ss1, 9, True)
    
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").Row = 0
    sc1.Item("Spread").Text = "◎"

    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 11, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 12, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 16, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 17, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   
   For i = 1 To ss1.MaxCols
      Call Gp_Sp_ColColor(ss1, i, , &H8000000B)
   Next i
    
   'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AKP1014C.P_SREFER2", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc2, Key:="Sc2"
    
    Call Gp_Sp_ColHidden(ss2, 1, True)
    Call Gp_Sp_ColHidden(ss2, 6, True)
    Call Gp_Sp_ColHidden(ss2, 7, True)
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
   
End Sub

Private Sub Cmd_Edit_Click()

    Dim i, j As Integer
    
    If Trim(txt_from_DATE.RawData) = "" Or Trim(txt_to_DATE.RawData) = "" Then
       MsgBox "查询日期未输入!", vbCritical, "系统提示信息"
       Exit Sub
    End If
    
    Call Gf_Sp_Display(M_CN1, ss2, Gf_Ms_MakeQuery(Proc_Sc("Sc2").Item("P-R"), "R", Mc1("pControl")))
    Call Gp_Sp_EvenRowBackcolor(sc2.Item("Spread"), 1)
    
    With ss2
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            Select Case .Text
            Case "A"
               .Col = 0
               .Text = "甲"
            Case "B"
               .Col = 0
               .Text = "乙"
            Case "C"
               .Col = 0
               .Text = "丙"
            Case "D"
               .Col = 0
               .Text = "丁"
            Case ""
               .Col = 0
               .Text = "合计"
               Call Gp_Sp_BlockColor(sc2.Item("Spread"), 1, .MaxCols, .MaxRows, .MaxRows, BLACK, &HE6E6FF)
            End Select
            
            For j = 2 To .MaxCols
                .Col = j
                If Val(.Text) = 0 Then
                   .Text = ""
                End If
            Next j
        Next i
    End With
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()
    'FormType = "Sheet"
    Screen.MousePointer = vbHourglass
    
   
    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_ControlLock(Mc1("lControl"), True)
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    'Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    Call Sp_Setting(Proc_Sc("Sc")("Spread"))
    Call Sp_Setting(Proc_Sc("Sc2")("Spread"))
    
    Call Gp_Spl_SizeGet(SSSplitter1, "K-System.INI", Me.Name, "H")
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc2")("Spread"), "K-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault
   
End Sub

Public Sub Form_Cls()

    Dim i, j As Integer
    Call Gf_Sp_Cls(sc1)
    'Call Gf_Sp_Cls(Sc2)
    '    If Gf_Sp_Cls(Sc1) Then
    '        Call Gp_Ms_Cls(Mc1("rControl"))
    '        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    '    End If
    With ss2
         For i = 1 To 12
             .Row = i
             For j = 1 To .MaxCols
                 .Col = j
                 .Text = ""
             Next j
         Next i
    End With
    
End Sub

Public Sub Form_Ref()

    Dim prod_wgt, upgrade_wgt, nograde_wgt As Double
    Dim ak, fl, jp, lw, hlw, qp As Double
    Dim i, j As Integer
    
    If Trim(txt_from_DATE.RawData) = "" Or Trim(txt_to_DATE.RawData) = "" Then
       MsgBox "查询日期未输入!", vbCritical, "系统提示信息"
       Exit Sub
    End If
    
        If Gf_Sp_Display(M_CN1, ss1, Gf_Ms_MakeQuery(Proc_Sc("Sc").Item("P-R"), "R", Mc1("pControl"))) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call Gp_Sp_EvenRowBackcolor(sc1.Item("Spread"), 1)
        End If
        
    With ss1
         For i = 1 To .MaxRows
             .Row = i
             For j = 2 To .MaxCols
                 .Col = j
                 If j <> 7 And Val(.Text) = 0 Then
                    .Text = ""
                 End If
             Next j
             If i = .MaxRows Then
                .Col = 0
                .Text = "合计"
                Call Gp_Sp_BlockColor(sc1.Item("Spread"), 1, .MaxCols, .MaxRows, .MaxRows, BLACK, &HE6E6FF)
             End If
             
         Next i
    End With
    
End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

   If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Spl_SizeSet(SSSplitter1, "K-System.INI", Me.Name)
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc2")("Spread"), "K-System.INI", Me.Name)
    
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
    
    Set iColumn2 = Nothing
    Set pColumn2 = Nothing
    Set lColumn2 = Nothing
    Set nColumn2 = Nothing
    Set mColumn2 = Nothing
    Set aColumn2 = Nothing
    
    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set sc2 = Nothing
    
    Set Proc_Sc = Nothing
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    'Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Form_Exc()

    ss1.Col = 0
    ss1.Row = 0
    If ss1.Text = "◎" Then
        Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    End If
    ss2.Col = 0
    ss2.Row = 0
    If ss2.Text = "◎" Then
        Call Gp_Sp_Excel(Me, Proc_Sc("Sc2")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    End If

End Sub

Public Sub Sp_Setting(ByVal sPname As Variant)

Dim iRow As Integer

  With sPname
'
  .RowHeight(-1) = 13
        
'        If .ColHeaderRows > 1 Then
'            .RowHeight(SpreadHeader + (.ColHeaderRows - 2)) = 13
'            .RowHeight(SpreadHeader + (.ColHeaderRows - 1)) = 13
'        Else
'            .RowHeight(0) = 24
'        End If
       
        .BackColorStyle = BackColorStyleUnderGrid
        
        .GrayAreaBackColor = &HE0E0E0
        .GridColor = &H808040
        
        .ShadowColor = &HE1E4CD
        .ShadowDark = &H808040
        .SelBackColor = &HCEECFF     ''&HE3F4FF      ''&HFFFF80     '&H808040
     
       
        .OperationMode = OperationModeNormal
        .RetainSelBlock = True
        .UserResize = UserResizeColumns
        
        .ProcessTab = True
        .ScrollBarExtMode = True
        .ButtonDrawMode = 1
        .TabStop = False
        
        .Col = 0: .Col2 = -1
        .Row = 0: .Row2 = -1
        
        
        .BlockMode = True
        .FontBold = False
        .FontName = "SimSun"
        .FontSize = 10
        .BlockMode = False
        
        .Col = -1
        .Row = 0
        .FontBold = True
    End With
'    With ss3
'         .ColHeadersShow = False
'         .Col = 1
'         .Row = 5
'         .RowHeight(5) = 30
'    End With
    
End Sub

Private Sub opt_bof_date_Click(Value As Integer)
    opt_bof_date.ForeColor = &HFF&
    opt_cc_date.ForeColor = &H808080
    TXT_BOF_CC.Text = "1"
End Sub

Private Sub opt_cc_date_Click(Value As Integer)
    opt_cc_date.ForeColor = &HFF&
    opt_bof_date.ForeColor = &H808080
    TXT_BOF_CC.Text = "2"
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").Row = 0
    sc1.Item("Spread").Text = "◎"

    sc2.Item("Spread").Col = 0
    sc2.Item("Spread").Row = 0
    sc2.Item("Spread").Text = ""
End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").Row = 0
    sc1.Item("Spread").Text = ""
    
    sc2.Item("Spread").Col = 0
    sc2.Item("Spread").Row = 0
    sc2.Item("Spread").Text = "◎"
End Sub
