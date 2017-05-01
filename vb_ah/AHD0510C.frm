VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form AHD0510C 
   Caption         =   "物流中心产成品转库报表_AHD0510C"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9225
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   16272
      _Version        =   196609
      SplitterBarWidth=   3
      BorderStyle     =   1
      PaneTree        =   "AHD0510C.frx":0000
      Begin Threed.SSFrame Single 
         Height          =   615
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   15105
         _ExtentX        =   26644
         _ExtentY        =   1085
         _Version        =   196609
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin InDate.ULabel ULabel5 
            Height          =   315
            Left            =   495
            Top             =   150
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
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
         Begin InDate.UDate DTP_DATE 
            Height          =   315
            Left            =   3360
            TabIndex        =   2
            Tag             =   "出库日期"
            Top             =   150
            Width           =   1455
            _ExtentX        =   2566
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
         Begin InDate.UDate DTP_DATE_FR 
            Height          =   315
            Left            =   1890
            TabIndex        =   6
            Tag             =   "出库日期"
            Top             =   150
            Width           =   1455
            _ExtentX        =   2566
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
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   8520
         Left            =   15
         TabIndex        =   3
         Top             =   690
         Width           =   15105
         _ExtentX        =   26644
         _ExtentY        =   15028
         _Version        =   393216
         Tab             =   2
         TabHeight       =   520
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
         TabCaption(0)   =   "产成品转库报表"
         TabPicture(0)   =   "AHD0510C.frx":0052
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "ss1"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "板卷、中板转出库"
         TabPicture(1)   =   "AHD0510C.frx":006E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "ss2"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "销售公司转库明细"
         TabPicture(2)   =   "AHD0510C.frx":008A
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "ss3"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin FPSpread.vaSpread ss3 
            Height          =   8220
            Left            =   0
            TabIndex        =   7
            Top             =   300
            Width           =   15105
            _Version        =   393216
            _ExtentX        =   26644
            _ExtentY        =   14499
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
            MaxCols         =   6
            MaxRows         =   20
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "AHD0510C.frx":00A6
         End
         Begin FPSpread.vaSpread ss1 
            Height          =   8220
            Left            =   -75000
            TabIndex        =   4
            Top             =   300
            Width           =   15105
            _Version        =   393216
            _ExtentX        =   26644
            _ExtentY        =   14499
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
            MaxCols         =   16
            MaxRows         =   20
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "AHD0510C.frx":066A
         End
         Begin FPSpread.vaSpread ss2 
            Height          =   8220
            Left            =   -75000
            TabIndex        =   5
            Top             =   300
            Width           =   15105
            _Version        =   393216
            _ExtentX        =   26644
            _ExtentY        =   14499
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
            MaxRows         =   20
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "AHD0510C.frx":1227
         End
      End
   End
End
Attribute VB_Name = "AHD0510C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       Nisco Production Management System
'-- Sub_System Name   Sale System
'-- Program Name      物流中心产成品转库报表
'-- Program ID        AHD0510C
'-- Document No       Q-00-0010(Specification)
'-- Designer          杨猛
'-- Coder             杨猛
'-- Date              2010.05.28
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE        EDITOR       DESCRIPTION
'-- 1.01  2010.05.28  杨猛
'-- 1.02  2010.11.01  李骞         销售公司转库明细
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------

Public FormType As String           'Form Type
Public Toolbar_St As String         'Active Form ToolBar Setting
Public sAuthority As String         'Active Form Authority Setting
Public sDateTime As String          'Active Form Time Setting
Public sQuery_load As String        'Active Form sQuery Setting

Dim pControl As New Collection      'Master Primary Key Collection
Dim nControl As New Collection      'Master Necessary Collection
Dim mControl As New Collection      'Master Maxlength check Collection
Dim iControl As New Collection      'Master Insert Collection
Dim rControl As New Collection      'Master Refer Collection
Dim cControl As New Collection      'Master Copy Collection
Dim aControl As New Collection      'Master -> Spread Collection
Dim lControl As New Collection      'Master Lock Collection

Dim pColumn  As New Collection      'Spread Primary Key Collection
Dim nColumn  As New Collection      'Spread necessary Column Collection
Dim mColumn  As New Collection      'Spread Maxlength check Column Collection
Dim iColumn  As New Collection      'Spread Insert Column Collection
Dim aColumn  As New Collection      'Master -> Spread Column Collection
Dim lColumn  As New Collection      'Spread Lock Column Collection

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
Dim Sc2 As New Collection           'Spread Collection
Dim Sc3 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Const SPD_C_D = 3
Const SPD_C_M = 4
Const SPD_F_DR = 5
Const SPD_F_MR = 6
Const SPD_F_DT = 7
Const SPD_F_MT = 8
Const SPD_F_DA = 9
Const SPD_F_MA = 10
Const SPD_T_DR = 11
Const SPD_T_MR = 12
Const SPD_T_DT = 13
Const SPD_T_MT = 14
Const SPD_T_DA = 15
Const SPD_T_MA = 16

Const SPD2_F_00 = 2
Const SPD2_F_ZB = 3
Const SPD2_F_52 = 4
Const SPD2_F_WD = 5
Const SPD2_F_SUM = 6

Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Refer"

     'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
  Call Gp_Ms_Collection(DTP_DATE_FR, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(DTP_DATE, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"

     Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
   
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AHD0510C.P_SREFER", Key:="P-R"
    sc1.Add Item:=pColumn, Key:="pColumn"
    sc1.Add Item:=nColumn, Key:="nColumn"
    sc1.Add Item:=aColumn, Key:="aColumn"
    sc1.Add Item:=mColumn, Key:="mColumn"
    sc1.Add Item:=iColumn, Key:="iColumn"
    sc1.Add Item:=lColumn, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)

    'Spread_Collection
    Sc2.Add Item:=ss2, Key:="Spread"
    Sc2.Add Item:="AHD0510C.P_SREFER1", Key:="P-R"
    Sc2.Add Item:=pColumn1, Key:="pColumn"
    Sc2.Add Item:=nColumn1, Key:="nColumn"
    Sc2.Add Item:=aColumn1, Key:="aColumn"
    Sc2.Add Item:=mColumn1, Key:="mColumn"
    Sc2.Add Item:=iColumn1, Key:="iColumn"
    Sc2.Add Item:=lColumn1, Key:="lColumn"
    Sc2.Add Item:=1, Key:="First"
    Sc2.Add Item:=ss2.MaxCols, Key:="Last"
    
    Call Gp_Sp_Collection(ss3, 1, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss3, 2, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss3, 3, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss3, 4, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss3, 5, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss3, 6, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    
    'Spread_Collection
    Sc3.Add Item:=ss3, Key:="Spread"
    Sc3.Add Item:="AHD0510C.P_SREFER2", Key:="P-R"
    Sc3.Add Item:=pColumn2, Key:="pColumn"
    Sc3.Add Item:=nColumn2, Key:="nColumn"
    Sc3.Add Item:=aColumn2, Key:="aColumn"
    Sc3.Add Item:=mColumn2, Key:="mColumn"
    Sc3.Add Item:=iColumn2, Key:="iColumn"
    Sc3.Add Item:=lColumn2, Key:="lColumn"
    Sc3.Add Item:=1, Key:="First"
    Sc3.Add Item:=ss3.MaxCols, Key:="Last"

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

    Call Gp_Ms_ControlLock(Mc1("lControl"), True)

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    Call Gp_Sp_Setting(ss2)
    Call Gp_Sp_Setting(ss3)
    
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    Call Gf_Sp_Cls(Sc2)
    Call Gf_Sp_Cls(Sc3)
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "H-System.INI", Me.Name)
    Call Gp_Sp_ColGet(ss2, "H-System.INI", Me.Name)
    Call Gp_Sp_ColGet(ss3, "H-System.INI", Me.Name)
    
    DTP_DATE_FR.RawData = Mid(DTP_DATE_FR.RawData, 1, 6) & "01"
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "H-System.INI", Me.Name)
    Call Gp_Sp_ColSet(ss2, "H-System.INI", Me.Name)
    Call Gp_Sp_ColSet(ss3, "H-System.INI", Me.Name)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set iColumn = Nothing
    Set pColumn = Nothing
    Set lColumn = Nothing
    Set nColumn = Nothing
    Set mColumn = Nothing
    Set aColumn = Nothing
    
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
    Set Sc2 = Nothing
    Set Sc3 = Nothing
    Set Proc_Sc = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Exit()

    Unload Me

End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(sc1) And Gf_Sp_Cls(Sc2) And Gf_Sp_Cls(Sc3) Then
       Call Gp_Ms_Cls(Mc1("rControl"))
       DTP_DATE_FR.RawData = Mid(DTP_DATE_FR.RawData, 1, 6) & "01"
       Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
       Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    End If

End Sub

Public Sub Master_Cpy()

'    Call Gf_Ms_Copy(Mc1)

End Sub

Public Sub Master_Pst()

'     If Gf_Ms_Paste(M_CN1, Mc1) Then
'        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
'     End If

End Sub

Public Sub Form_Ref()
    
Dim iRow   As Integer
Dim iCol   As Integer
Dim iWgt00 As Double
Dim iWgtZB As Double
Dim iWgt52 As Double
Dim iWgtWD As Double

    If Not Gp_DateCheck(DTP_DATE.Text, "S") Then
       Call Gp_MsgBoxDisplay("请正确输入时间..")
       Exit Sub
    End If
                        
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    Select Case SSTab1.Tab
           
           Case 0
     
                If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
                    Call Data_Sum_Edit
                    ss1.OperationMode = OperationModeNormal
                    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                End If
            
           Case 1
           
                If Gf_Sp_Refer(M_CN1, Sc2, Mc1, Mc1("nControl"), Mc1("mControl")) Then
                    ss2.OperationMode = OperationModeNormal
                    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                End If
           Case 2
           
                If Gf_Sp_Refer(M_CN1, Sc3, Mc1, Mc1("nControl"), Mc1("mControl")) Then
                    ss3.OperationMode = OperationModeNormal
                    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                    For iRow = 1 To ss3.MaxRows
                        ss3.Row = iRow
                        ss3.Col = SPD2_F_00
                        iWgt00 = ss3.Value
                        ss3.Col = SPD2_F_ZB
                        iWgtZB = ss3.Value
                        ss3.Col = SPD2_F_52
                        iWgt52 = ss3.Value
                        ss3.Col = SPD2_F_WD
                        iWgtWD = ss3.Value
                        ss3.Col = SPD2_F_SUM
                        ss3.Text = iWgt00 + iWgtZB + iWgt52 + iWgtWD
                    Next iRow
                    ' 转库量为 0 时用 "" 替代
                    For iRow = 1 To ss3.MaxRows
                         ss3.Row = iRow
                         For iCol = SPD2_F_00 To SPD2_F_SUM
                             ss3.Col = iCol:
                             If ss3.Value = 0 Then ss3.Text = ""
                         Next iCol
                    Next iRow
                    
                End If

     End Select
               
End Sub

Public Sub Form_Pro()

'     If Gf_Mc_Authority(sAuthority, Mc1) Then
'       ' txt_ins_emp.Text = sUserID
'       If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
'    End If

End Sub

Public Sub Form_Del()

'    If Not Gf_Ms_Del(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

End Sub

Public Sub Form_Exc()
  
    Select Case SSTab1.Tab
           
           Case 0
           
                Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
            
           Case 1
           
                Call Gp_Sp_Excel(Me, ss2, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
                
           Case 2
           
                Call Gp_Sp_Excel(Me, ss3, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    
    End Select
    
End Sub

Private Sub Data_Sum_Edit()

    Dim iRow           As Integer
    Dim iCol           As Integer
    Dim iTwgt          As Double
    Dim iRwgt          As Double
    Dim iAwgt          As Double
    Dim iWgt           As Double
    Dim iTolWgt        As Double
    
    If ss1.MaxRows <= 0 Then Exit Sub
    
    iRow = 0
    
    With ss1
    
        ' 计算火车转库量
        ' 火车转库量 = 转库总量 - 汽车转库量
        For iRow = 1 To .MaxRows
        
            .Row = iRow
            
            .Col = SPD_F_DR:    iTwgt = .Value
            .Col = SPD_F_DA:    iAwgt = .Value
            .Col = SPD_F_DT:    iRwgt = iAwgt - iTwgt:     .Value = iRwgt
            
            .Col = SPD_F_MR:    iTwgt = .Value
            .Col = SPD_F_MA:    iAwgt = .Value
            .Col = SPD_F_MT:    iRwgt = iAwgt - iTwgt:     .Value = iRwgt
            
            .Col = SPD_T_DR:    iTwgt = .Value
            .Col = SPD_T_DA:    iAwgt = .Value
            .Col = SPD_T_DT:    iRwgt = iAwgt - iTwgt:     .Value = iRwgt
            
            .Col = SPD_T_MR:    iTwgt = .Value
            .Col = SPD_T_MA:    iAwgt = .Value
            .Col = SPD_T_MT:    iRwgt = iAwgt - iTwgt:     .Value = iRwgt
          
        Next iRow
        
        ' 转库汇总
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows:    .Col = 2:    .Text = "合   计"
    
        For iCol = SPD_C_D To .MaxCols
            iTolWgt = 0
            .Col = iCol
            For iRow = 1 To .MaxRows
                .Row = iRow
                 If iRow <> .MaxRows Then
                    iWgt = .Value
                    iTolWgt = iTolWgt + iWgt
                 Else
                    .Value = iTolWgt
                 End If
            Next iRow
        Next iCol
            
        ' 转库量为 0 时用 "" 替代
        For iRow = 1 To .MaxRows
        
            .Row = iRow
            
             For iCol = SPD_C_D To .MaxCols
                 .Col = iCol:       iAwgt = .Value
                 If iAwgt = 0 Then .Text = ""
             Next iCol
          
        Next iRow
        
        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ss1.Row, ss1.Row, , &HFFFFC0)
                
    End With

End Sub

Private Sub ExcelPrn()

    Dim I               As Integer
    Dim xlApp           As Object
    Dim xlSheet         As Object
    Dim sDateFr         As String
    Dim sDateTo         As String

    If ss1.MaxRows < 1 Then Exit Sub

    Screen.MousePointer = vbHourglass

    On Error Resume Next

    Set xlApp = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then
        Set xlApp = CreateObject("Excel.Application")
    End If

    Err.Clear

    xlApp.Workbooks.Open (App.Path & "\AGC2042C.xls")

    Set xlSheet = xlApp.Worksheets("Sheet1")
    xlApp.Sheets("Sheet1").Select
    
    For I = 2 To ss1.MaxRows
          xlApp.Rows("4:4").Select
          xlApp.Selection.Copy
          xlApp.Selection.Insert Shift:=1
    Next I

    xlApp.Range("B1").Value = Left(sDateFr, 4) + "年" + Mid(sDateFr, 6, 2) + "月" + Mid(sDateFr, 9, 2) + "日 - " _
                  + Left(sDateTo, 4) + "年" + Mid(sDateTo, 6, 2) + "月" + Mid(sDateTo, 9, 2) + "日 "

    Clipboard.Clear
    ss1.SetSelection 1, 1, ss1.MaxCols, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("A4").Select
    xlApp.ActiveSheet.Paste
    Clipboard.Clear

    xlApp.Range("I2").Select
    xlApp.ActiveSheet.Paste

'    xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True

    ss1.ClearSelection

    Screen.MousePointer = vbDefault

    xlApp.Application.Visible = True
'     xlApp.Application.Visible = False
'     xlSheet.Close False
'     xlApp.Quit

    Set xlSheet = Nothing
    Set xlApp = Nothing

    Exit Sub

ErrHandle:
    MsgBox Error
'    xlApp.Application.Visible = True

    Set xlSheet = Nothing
    Set xlApp = Nothing
    Screen.MousePointer = vbDefault
End Sub





