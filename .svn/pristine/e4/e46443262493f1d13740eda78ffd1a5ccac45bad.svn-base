VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form ACF0060C 
   Caption         =   "板材订单汇总表_ACF0060C"
   ClientHeight    =   9225
   ClientLeft      =   285
   ClientTop       =   2325
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15240
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8580
      Left            =   60
      TabIndex        =   0
      Top             =   600
      Width           =   15060
      _ExtentX        =   26564
      _ExtentY        =   15134
      _Version        =   196609
      SplitterBarWidth=   3
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "ACF0060C.frx":0000
      Begin FPSpread.vaSpread ss1 
         Height          =   8580
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   15060
         _Version        =   393216
         _ExtentX        =   26564
         _ExtentY        =   15134
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         ButtonDrawMode  =   4
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
         MaxRows         =   1
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         SpreadDesigner  =   "ACF0060C.frx":0032
      End
   End
   Begin InDate.ULabel ULabel3 
      DragMode        =   1  'Automatic
      Height          =   315
      Left            =   240
      Top             =   80
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "生产日期"
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
   Begin InDate.UDate prod_date_to 
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Tag             =   "生产日期"
      Top             =   75
      Width           =   1410
      _ExtentX        =   2487
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
   Begin VB.Line Line1 
      X1              =   120
      X2              =   15120
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   15105
      Y1              =   480
      Y2              =   480
   End
End
Attribute VB_Name = "ACF0060C"
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
'-- Program ID        ACB1022C
'-- Document No       Q-00-0010(Specification)
'-- Designer          HJD
'-- Coder             HJD
'-- Date              2003.9.26
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

Dim pContro1 As New Collection      'Master Primary Key Collection
Dim nContro1 As New Collection      'Master Necessary Collection
Dim mContro1 As New Collection      'Master Maxlength check Collection
Dim iContro1 As New Collection      'Master Insert Collection
Dim rContro1 As New Collection      'Master Refer Collection
Dim cContro1 As New Collection      'Master Copy Collection
Dim aContro1 As New Collection      'Master -> Spread Collection
Dim lContro1 As New Collection      'Master Lock Collection

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

'Const SS2_PLT = 1

Const SS1_PLAN_WGT = 3
Const SS1_ACT_WGT = 4
Const SS1_OWE_WGT = 5
Const SS1_DECLARE_RATIO = 6
Const SS1_N_WGT = 7
Const SS1_T_WGT = 8
Const SS1_QT_WGT = 9
Const SS1_NT_WGT = 10
Const SS1_TOTAL_WGT = 11



Dim sWgtLenFlag As String
Dim sQuery  As String

Private Sub Form_Define()


        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
     Call Gp_Ms_Collection(prod_date_to, "p", "n", "", " ", "r", " ", "", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
     
    'MASTER Collection
    Mc1.Add Item:=pContro1, Key:="pControl"
    Mc1.Add Item:=nContro1, Key:="nControl"
    Mc1.Add Item:=mContro1, Key:="mControl"
    Mc1.Add Item:=iContro1, Key:="iControl"
    Mc1.Add Item:=rContro1, Key:="rControl"
    Mc1.Add Item:=cContro1, Key:="cControl"
    Mc1.Add Item:=aContro1, Key:="aControl"
    Mc1.Add Item:=lContro1, Key:="lControl"
    

    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss1, 1, "", " ", " ", "", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", "i ", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", "", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", "", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", "", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", "", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", "i", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    

    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="ACF0060C.P_REFER1", Key:="P-R"
    sc1.Add Item:="ACF0060C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="ACF0060C.P_ONEROW", Key:="P-O"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc1, Key:="Sc1"
    
    Call Gp_Sp_ColHidden(ss1, 12, True)
    
    
'    sc1.Item("Spread").Col = 0
'    sc1.Item("Spread").Row = 0
'    sc1.Item("Spread").Text = "◎"
 
'
   Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    
        
End Sub



Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
    MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Insert
    MDIMain.MenuTool.Buttons(8).Enabled = False    'Row Delete
    MDIMain.MenuTool.Buttons(9).Enabled = False    'Row Cancle
    MDIMain.MenuTool.Buttons(11).Enabled = False   'Copy
    MDIMain.MenuTool.Buttons(12).Enabled = False   'Paste

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Public Sub Form_Pro()
   
    If Gf_Sp_Process(M_CN1, Proc_Sc("Sc1"), Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
    
    Call Form_Ref
    
End Sub


Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Insert
    MDIMain.MenuTool.Buttons(8).Enabled = False    'Row Delete
    MDIMain.MenuTool.Buttons(9).Enabled = False    'Row Cancle
    MDIMain.MenuTool.Buttons(11).Enabled = False   'Copy
    MDIMain.MenuTool.Buttons(12).Enabled = False   'Paste
    
    Call Gp_Ms_Cls(Mc1("rControl"))

    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)

    

    
    Call Gf_Sp_Cls(sc1)

    
    Call Gp_Spl_SizeGet(SSSplitter1, "C-System.INI", Me.Name, "H")
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "C-System.INI", Me.Name)
    
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc1")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Spl_SizeSet(SSSplitter1, "C-System.INI", Me.Name)
    
    Call Gp_Sp_ColSet(sc1.Item("Spread"), "C-System.INI", Me.Name)
    
    Set pContro1 = Nothing
    Set nContro1 = Nothing
    Set iContro1 = Nothing
    Set rContro1 = Nothing
    Set cContro1 = Nothing
    Set aContro1 = Nothing
    Set lContro1 = Nothing
    Set mContro1 = Nothing
      
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

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

Public Sub Form_Cls()
    
        If Gf_Sp_Cls(sc1) Then
            Call Gp_Ms_Cls(Mc1("rControl"))
            Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
            MDIMain.MenuTool.Buttons(7).Enabled = False
            MDIMain.MenuTool.Buttons(8).Enabled = False
            MDIMain.MenuTool.Buttons(9).Enabled = False
            MDIMain.MenuTool.Buttons(11).Enabled = False
            MDIMain.MenuTool.Buttons(12).Enabled = False
            Call Gp_Ms_ControlLock(Mc1("lControl"), False)
            rContro1(1).SetFocus
   
        End If
    
End Sub

Public Sub Form_Ref()

Dim PLAN_WGT_RE        As Double
Dim ACT_WGT_RE         As Double
Dim OWE_WGT_RE         As Double
Dim DECLARE_RATIO_RE   As Double
Dim N_WGT_RE           As Double
Dim T_WGT_RE           As Double
Dim QT_WGT_RE           As Double
Dim NT_WGT_RE           As Double
Dim TOTAL_HTM_WGT_RE        As Double
Dim PLAN_WGT_EX        As Double
Dim ACT_WGT_EX         As Double
Dim OWE_WGT_EX         As Double
Dim DECLARE_RATIO_EX   As Double
Dim N_WGT_EX           As Double
Dim T_WGT_EX           As Double
Dim QT_WGT_EX           As Double
Dim NT_WGT_EX           As Double
Dim TOTAL_HTM_WGT_EX        As Double
Dim PLAN_WGT_ALL        As Double
Dim ACT_WGT_ALL         As Double
Dim OWE_WGT_ALL         As Double
Dim DECLARE_RATIO_ALL   As Double
Dim N_WGT_ALL           As Double
Dim T_WGT_ALL           As Double
Dim QT_WGT_ALL           As Double
Dim NT_WGT_ALL           As Double
Dim TOTAL_HTM_WGT_ALL        As Double
Dim i                   As Integer

         Call Gf_Sp_Cls(sc1)
            Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
            MDIMain.MenuTool.Buttons(7).Enabled = False
            MDIMain.MenuTool.Buttons(8).Enabled = False
            MDIMain.MenuTool.Buttons(9).Enabled = False
            MDIMain.MenuTool.Buttons(11).Enabled = False
            MDIMain.MenuTool.Buttons(12).Enabled = False
            Call Gp_Ms_ControlLock(Mc1("lControl"), False)
            rContro1(1).SetFocus
   
       
    

    
    'If Gf_Sp_ProceExist(sc2.Item("Spread")) Then Exit Sub
    
    Call Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl"))
        ss1.OperationMode = OperationModeNormal
        'Call Gp_Sp_BlockColor(ss1, 7, 7, 1, ss1.MaxRows)
        
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        
        MDIMain.MenuTool.Buttons(4).Enabled = True
        MDIMain.MenuTool.Buttons(7).Enabled = False
        MDIMain.MenuTool.Buttons(8).Enabled = False
        MDIMain.MenuTool.Buttons(9).Enabled = False
        MDIMain.MenuTool.Buttons(11).Enabled = False
        MDIMain.MenuTool.Buttons(12).Enabled = False
        
        
        
        
        With ss1
        If .MaxRows < 1 Then
           Exit Sub
        End If
        
        .MaxRows = .MaxRows + 1
        .Row = 1
        
         '研销计划量
        
         .Col = SS1_PLAN_WGT:                PLAN_WGT_RE = Val(.Text)
         '研销申报量
         .Col = SS1_ACT_WGT:                 ACT_WGT_RE = Val(.Text)
         '超欠
          OWE_WGT_RE = ACT_WGT_RE - PLAN_WGT_RE
          '申报比例
          If PLAN_WGT_RE = 0 Then
          DECLARE_RATIO_RE = 0
          Else
          DECLARE_RATIO_RE = ACT_WGT_RE / PLAN_WGT_RE
          End If
          
          '热处理计算
          .Col = SS1_N_WGT:                 N_WGT_RE = Val(.Text)
          .Col = SS1_T_WGT:                 T_WGT_RE = Val(.Text)
          .Col = SS1_QT_WGT:                QT_WGT_RE = Val(.Text)
          .Col = SS1_NT_WGT:                NT_WGT_RE = Val(.Text)
          
          TOTAL_HTM_WGT_RE = N_WGT_RE + T_WGT_RE + QT_WGT_RE + NT_WGT_RE
          
          .Col = SS1_OWE_WGT:                .Text = OWE_WGT_RE
          .Col = SS1_DECLARE_RATIO:                .Text = Round(DECLARE_RATIO_RE, 2)
          .Col = SS1_TOTAL_WGT:                .Text = TOTAL_HTM_WGT_RE
          
          .Row = 2
        
         '计划量
        
         .Col = SS1_PLAN_WGT:                PLAN_WGT_EX = Val(.Text)
         '申报量
         .Col = SS1_ACT_WGT:                 ACT_WGT_EX = Val(.Text)
         '超欠
          OWE_WGT_EX = ACT_WGT_EX - PLAN_WGT_EX
          '申报比例
          If PLAN_WGT_EX = 0 Then
          DECLARE_RATIO_EX = 0
          Else
          DECLARE_RATIO_EX = ACT_WGT_EX / PLAN_WGT_EX
          End If
          
          '热处理计算
          .Col = SS1_N_WGT:                 N_WGT_EX = Val(.Text)
          .Col = SS1_T_WGT:                 T_WGT_EX = Val(.Text)
          .Col = SS1_QT_WGT:                QT_WGT_EX = Val(.Text)
          .Col = SS1_NT_WGT:                NT_WGT_EX = Val(.Text)
          
          TOTAL_HTM_WGT_EX = N_WGT_EX + T_WGT_EX + QT_WGT_EX + NT_WGT_EX
          
          .Col = SS1_OWE_WGT:                .Text = OWE_WGT_EX
          .Col = SS1_DECLARE_RATIO:                .Text = Round(DECLARE_RATIO_EX, 2)
          .Col = SS1_TOTAL_WGT:                .Text = TOTAL_HTM_WGT_EX
          
          
          .Row = 3
          PLAN_WGT_ALL = PLAN_WGT_RE + PLAN_WGT_EX
          ACT_WGT_ALL = ACT_WGT_RE + ACT_WGT_EX
          OWE_WGT_ALL = ACT_WGT_ALL - PLAN_WGT_ALL
          
          If PLAN_WGT_ALL = 0 Then
          
          DECLARE_RATIO_ALL = 0
          
          Else
          
          DECLARE_RATIO_ALL = ACT_WGT_ALL / PLAN_WGT_ALL
          
          End If
          
          
          N_WGT_ALL = N_WGT_RE + N_WGT_EX
          T_WGT_ALL = T_WGT_RE + T_WGT_EX
          QT_WGT_ALL = QT_WGT_RE + QT_WGT_EX
          NT_WGT_ALL = NT_WGT_RE + NT_WGT_EX
          TOTAL_HTM_WGT_ALL = TOTAL_HTM_WGT_EX + TOTAL_HTM_WGT_RE
          
          
          
         .Col = SS1_PLAN_WGT:                  .Text = PLAN_WGT_ALL
         .Col = SS1_ACT_WGT:                   .Text = ACT_WGT_ALL
         .Col = SS1_OWE_WGT:                   .Text = OWE_WGT_ALL
         .Col = SS1_DECLARE_RATIO:             .Text = Round(DECLARE_RATIO_ALL, 2)
         .Col = SS1_N_WGT:                     .Text = N_WGT_ALL
         .Col = SS1_T_WGT:                     .Text = T_WGT_ALL
         .Col = SS1_QT_WGT:                    .Text = QT_WGT_ALL
         .Col = SS1_NT_WGT:                    .Text = NT_WGT_ALL
         .Col = SS1_TOTAL_WGT:                 .Text = TOTAL_HTM_WGT_ALL
         
         
        For i = 1 To ss1.MaxCols
               ss1.Col = i
               ss1.BackColor = &HC0C0FF
          Next
         
         
         .Row = 4
        
         '研销计划量
        
         .Col = SS1_PLAN_WGT:                PLAN_WGT_RE = Val(.Text)
         '研销申报量
         .Col = SS1_ACT_WGT:                 ACT_WGT_RE = Val(.Text)
         '超欠
          OWE_WGT_RE = ACT_WGT_RE - PLAN_WGT_RE
          '申报比例
          If PLAN_WGT_RE = 0 Then
          DECLARE_RATIO_RE = 0
          Else
          DECLARE_RATIO_RE = ACT_WGT_RE / PLAN_WGT_RE
          End If
          
          '热处理计算
          .Col = SS1_N_WGT:                 N_WGT_RE = Val(.Text)
          .Col = SS1_T_WGT:                 T_WGT_RE = Val(.Text)
          .Col = SS1_QT_WGT:                QT_WGT_RE = Val(.Text)
          .Col = SS1_NT_WGT:                NT_WGT_RE = Val(.Text)
          
          TOTAL_HTM_WGT_RE = N_WGT_RE + T_WGT_RE + QT_WGT_RE + NT_WGT_RE
          
          .Col = SS1_OWE_WGT:                .Text = OWE_WGT_RE
          .Col = SS1_DECLARE_RATIO:                .Text = Round(DECLARE_RATIO_RE, 2)
          .Col = SS1_TOTAL_WGT:                .Text = TOTAL_HTM_WGT_RE
          
          .Row = 5
        
         '计划量
        
         .Col = SS1_PLAN_WGT:                PLAN_WGT_EX = Val(.Text)
         '申报量
         .Col = SS1_ACT_WGT:                 ACT_WGT_EX = Val(.Text)
         '超欠
          OWE_WGT_EX = ACT_WGT_EX - PLAN_WGT_EX
          '申报比例
          If PLAN_WGT_EX = 0 Then
          DECLARE_RATIO_EX = 0
          Else
          DECLARE_RATIO_EX = ACT_WGT_EX / PLAN_WGT_EX
          End If
          
          '热处理计算
          .Col = SS1_N_WGT:                 N_WGT_EX = Val(.Text)
          .Col = SS1_T_WGT:                 T_WGT_EX = Val(.Text)
          .Col = SS1_QT_WGT:                QT_WGT_EX = Val(.Text)
          .Col = SS1_NT_WGT:                NT_WGT_EX = Val(.Text)
          
          TOTAL_HTM_WGT_EX = N_WGT_EX + T_WGT_EX + QT_WGT_EX + NT_WGT_EX
          
          .Col = SS1_OWE_WGT:                .Text = OWE_WGT_EX
          .Col = SS1_DECLARE_RATIO:                .Text = Round(DECLARE_RATIO_EX, 2)
          .Col = SS1_TOTAL_WGT:                .Text = TOTAL_HTM_WGT_EX
          
          
          .Row = 6
          PLAN_WGT_ALL = PLAN_WGT_RE + PLAN_WGT_EX
          ACT_WGT_ALL = ACT_WGT_RE + ACT_WGT_EX
          OWE_WGT_ALL = ACT_WGT_ALL - PLAN_WGT_ALL
          
          If PLAN_WGT_ALL = 0 Then
          
          DECLARE_RATIO_ALL = 0
          
          Else
          
          DECLARE_RATIO_ALL = ACT_WGT_ALL / PLAN_WGT_ALL
          
          End If
          
          N_WGT_ALL = N_WGT_RE + N_WGT_EX
          T_WGT_ALL = T_WGT_RE + T_WGT_EX
          QT_WGT_ALL = QT_WGT_RE + QT_WGT_EX
          NT_WGT_ALL = NT_WGT_RE + NT_WGT_EX
          TOTAL_HTM_WGT_ALL = TOTAL_HTM_WGT_EX + TOTAL_HTM_WGT_RE
          
          
          
         .Col = SS1_PLAN_WGT:                  .Text = PLAN_WGT_ALL
         .Col = SS1_ACT_WGT:                   .Text = ACT_WGT_ALL
         .Col = SS1_OWE_WGT:                   .Text = OWE_WGT_ALL
         .Col = SS1_DECLARE_RATIO:             .Text = Round(DECLARE_RATIO_ALL, 2)
         .Col = SS1_N_WGT:                     .Text = N_WGT_ALL
         .Col = SS1_T_WGT:                     .Text = T_WGT_ALL
         .Col = SS1_QT_WGT:                    .Text = QT_WGT_ALL
         .Col = SS1_NT_WGT:                    .Text = NT_WGT_ALL
         .Col = SS1_TOTAL_WGT:                 .Text = TOTAL_HTM_WGT_ALL
         
         
         For i = 1 To ss1.MaxCols
               ss1.Col = i
               ss1.BackColor = &HC0C0FF
          Next
         
         
         
         .Row = 7
        
         '研销计划量
        
         .Col = SS1_PLAN_WGT:                PLAN_WGT_RE = Val(.Text)
         '研销申报量
         .Col = SS1_ACT_WGT:                 ACT_WGT_RE = Val(.Text)
         '超欠
          OWE_WGT_RE = ACT_WGT_RE - PLAN_WGT_RE
          '申报比例
          If PLAN_WGT_RE = 0 Then
          DECLARE_RATIO_RE = 0
          Else
          DECLARE_RATIO_RE = ACT_WGT_RE / PLAN_WGT_RE
          End If
          
          '热处理计算
          .Col = SS1_N_WGT:                 N_WGT_RE = Val(.Text)
          .Col = SS1_T_WGT:                 T_WGT_RE = Val(.Text)
          .Col = SS1_QT_WGT:                QT_WGT_RE = Val(.Text)
          .Col = SS1_NT_WGT:                NT_WGT_RE = Val(.Text)
          
          TOTAL_HTM_WGT_RE = N_WGT_RE + T_WGT_RE + QT_WGT_RE + NT_WGT_RE
          
          .Col = SS1_OWE_WGT:                .Text = OWE_WGT_RE
          .Col = SS1_DECLARE_RATIO:                .Text = Round(DECLARE_RATIO_RE, 2)
          .Col = SS1_TOTAL_WGT:                .Text = TOTAL_HTM_WGT_RE
          
          .Row = 8
        
         '计划量
        
         .Col = SS1_PLAN_WGT:                PLAN_WGT_EX = Val(.Text)
         '申报量
         .Col = SS1_ACT_WGT:                 ACT_WGT_EX = Val(.Text)
         '超欠
          OWE_WGT_EX = ACT_WGT_EX - PLAN_WGT_EX
          '申报比例
          If PLAN_WGT_EX = 0 Then
          DECLARE_RATIO_EX = 0
          Else
          DECLARE_RATIO_EX = ACT_WGT_EX / PLAN_WGT_EX
          End If
          
          '热处理计算
          .Col = SS1_N_WGT:                 N_WGT_EX = Val(.Text)
          .Col = SS1_T_WGT:                 T_WGT_EX = Val(.Text)
          .Col = SS1_QT_WGT:                QT_WGT_EX = Val(.Text)
          .Col = SS1_NT_WGT:                NT_WGT_EX = Val(.Text)
          
          TOTAL_HTM_WGT_EX = N_WGT_EX + T_WGT_EX + QT_WGT_EX + NT_WGT_EX
          
          .Col = SS1_OWE_WGT:                .Text = OWE_WGT_EX
          .Col = SS1_DECLARE_RATIO:                .Text = Round(DECLARE_RATIO_EX, 2)
          .Col = SS1_TOTAL_WGT:                .Text = TOTAL_HTM_WGT_EX
          
          
          .Row = 9
          PLAN_WGT_ALL = PLAN_WGT_RE + PLAN_WGT_EX
          ACT_WGT_ALL = ACT_WGT_RE + ACT_WGT_EX
          OWE_WGT_ALL = ACT_WGT_ALL - PLAN_WGT_ALL
          
          If PLAN_WGT_ALL = 0 Then
          
          DECLARE_RATIO_ALL = 0
          
          Else
          
          DECLARE_RATIO_ALL = ACT_WGT_ALL / PLAN_WGT_ALL
          
          End If
          
          N_WGT_ALL = N_WGT_RE + N_WGT_EX
          T_WGT_ALL = T_WGT_RE + T_WGT_EX
          QT_WGT_ALL = QT_WGT_RE + QT_WGT_EX
          NT_WGT_ALL = NT_WGT_RE + NT_WGT_EX
          TOTAL_HTM_WGT_ALL = TOTAL_HTM_WGT_EX + TOTAL_HTM_WGT_RE
          
          
         .Col = SS1_PLAN_WGT:                  .Text = PLAN_WGT_ALL
         .Col = SS1_ACT_WGT:                   .Text = ACT_WGT_ALL
         .Col = SS1_OWE_WGT:                   .Text = OWE_WGT_ALL
         .Col = SS1_DECLARE_RATIO:             .Text = Round(DECLARE_RATIO_ALL, 2)
         .Col = SS1_N_WGT:                     .Text = N_WGT_ALL
         .Col = SS1_T_WGT:                     .Text = T_WGT_ALL
         .Col = SS1_QT_WGT:                    .Text = QT_WGT_ALL
         .Col = SS1_NT_WGT:                    .Text = NT_WGT_ALL
         .Col = SS1_TOTAL_WGT:                 .Text = TOTAL_HTM_WGT_ALL
         
         
         For i = 1 To ss1.MaxCols
               ss1.Col = i
               ss1.BackColor = &HC0C0FF
          Next
         
         
         .Row = 10
        
         '研销计划量
        
         .Col = SS1_PLAN_WGT:                PLAN_WGT_RE = Val(.Text)
         '研销申报量
         .Col = SS1_ACT_WGT:                 ACT_WGT_RE = Val(.Text)
         '超欠
          OWE_WGT_RE = ACT_WGT_RE - PLAN_WGT_RE
          '申报比例
          If PLAN_WGT_RE = 0 Then
          DECLARE_RATIO_RE = 0
          Else
          DECLARE_RATIO_RE = ACT_WGT_RE / PLAN_WGT_RE
          End If
          
          '热处理计算
          .Col = SS1_N_WGT:                 N_WGT_RE = Val(.Text)
          .Col = SS1_T_WGT:                 T_WGT_RE = Val(.Text)
          .Col = SS1_QT_WGT:                QT_WGT_RE = Val(.Text)
          .Col = SS1_NT_WGT:                NT_WGT_RE = Val(.Text)
          
          TOTAL_HTM_WGT_RE = N_WGT_RE + T_WGT_RE + QT_WGT_RE + NT_WGT_RE
          
          .Col = SS1_OWE_WGT:                .Text = OWE_WGT_RE
          .Col = SS1_DECLARE_RATIO:                .Text = Round(DECLARE_RATIO_RE, 2)
          .Col = SS1_TOTAL_WGT:                .Text = TOTAL_HTM_WGT_RE
          
          .Row = 11
        
         '计划量
        
         .Col = SS1_PLAN_WGT:                PLAN_WGT_EX = Val(.Text)
         '申报量
         .Col = SS1_ACT_WGT:                 ACT_WGT_EX = Val(.Text)
         '超欠
          OWE_WGT_EX = ACT_WGT_EX - PLAN_WGT_EX
          '申报比例
          If PLAN_WGT_EX = 0 Then
          DECLARE_RATIO_EX = 0
          Else
          DECLARE_RATIO_EX = ACT_WGT_EX / PLAN_WGT_EX
          End If
          
          '热处理计算
          .Col = SS1_N_WGT:                 N_WGT_EX = Val(.Text)
          .Col = SS1_T_WGT:                 T_WGT_EX = Val(.Text)
          .Col = SS1_QT_WGT:                QT_WGT_EX = Val(.Text)
          .Col = SS1_NT_WGT:                NT_WGT_EX = Val(.Text)
          
          TOTAL_HTM_WGT_EX = N_WGT_EX + T_WGT_EX + QT_WGT_EX + NT_WGT_EX
          
          .Col = SS1_OWE_WGT:                .Text = OWE_WGT_EX
          .Col = SS1_DECLARE_RATIO:                .Text = Round(DECLARE_RATIO_EX, 2)
          .Col = SS1_TOTAL_WGT:                .Text = TOTAL_HTM_WGT_EX
          
          
          .Row = 12
          PLAN_WGT_ALL = PLAN_WGT_RE + PLAN_WGT_EX
          ACT_WGT_ALL = ACT_WGT_RE + ACT_WGT_EX
          OWE_WGT_ALL = ACT_WGT_ALL - PLAN_WGT_ALL
          
          If PLAN_WGT_ALL = 0 Then
          
          DECLARE_RATIO_ALL = 0
          
          Else
          
          DECLARE_RATIO_ALL = ACT_WGT_ALL / PLAN_WGT_ALL
          
          End If
          
          N_WGT_ALL = N_WGT_RE + N_WGT_EX
          T_WGT_ALL = T_WGT_RE + T_WGT_EX
          QT_WGT_ALL = QT_WGT_RE + QT_WGT_EX
          NT_WGT_ALL = NT_WGT_RE + NT_WGT_EX
          TOTAL_HTM_WGT_ALL = TOTAL_HTM_WGT_EX + TOTAL_HTM_WGT_RE
          
          
          
         .Col = SS1_PLAN_WGT:                  .Text = PLAN_WGT_ALL
         .Col = SS1_ACT_WGT:                   .Text = ACT_WGT_ALL
         .Col = SS1_OWE_WGT:                   .Text = OWE_WGT_ALL
         .Col = SS1_DECLARE_RATIO:             .Text = Round(DECLARE_RATIO_ALL, 2)
         .Col = SS1_N_WGT:                     .Text = N_WGT_ALL
         .Col = SS1_T_WGT:                     .Text = T_WGT_ALL
         .Col = SS1_QT_WGT:                    .Text = QT_WGT_ALL
         .Col = SS1_NT_WGT:                    .Text = NT_WGT_ALL
         .Col = SS1_TOTAL_WGT:                 .Text = TOTAL_HTM_WGT_ALL
         
         
         For i = 1 To ss1.MaxCols
               ss1.Col = i
               ss1.BackColor = &HC0C0FF
          Next
          
        
        End With
        
End Sub


Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1
    
End Sub

Public Sub Form_Exc()

        Call Gp_Sp_Excel(Me, Proc_Sc("Sc1")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub




Private Sub ss1_EditChange(ByVal Col As Long, ByVal Row As Long)
ss1.Row = Row
    
    ss1.Col = 0
    ss1.Row = Row
    Select Case Trim(ss1.Text)
          Case "Input", "Update", "Delete"
          Case Else
               ss1.Text = "Update"
    End Select
End Sub
