VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form CKP1020C 
   Caption         =   "中板厂订单完成情况简报_CKP1020C"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.TextBox ORD_OK 
      Alignment       =   2  'Center
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
      Left            =   15570
      MaxLength       =   1
      TabIndex        =   12
      Text            =   "A"
      Top             =   210
      Visible         =   0   'False
      Width           =   585
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9165
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   15210
      _ExtentX        =   26829
      _ExtentY        =   16166
      _Version        =   196609
      SplitterBarWidth=   3
      BorderStyle     =   0
      Locked          =   -1  'True
      PaneTree        =   "CKP1020C.frx":0000
      Begin FPSpread.vaSpread ss1 
         Height          =   8535
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   630
         Width           =   15210
         _Version        =   393216
         _ExtentX        =   26829
         _ExtentY        =   15055
         _StockProps     =   64
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
         MaxCols         =   10
         MaxRows         =   20
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         SpreadDesigner  =   "CKP1020C.frx":0052
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   570
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   15210
         _ExtentX        =   26829
         _ExtentY        =   1005
         _Version        =   196609
         BackColor       =   14737632
         Begin VB.OptionButton ORD_NO_N 
            BackColor       =   &H00E0E0E0&
            Caption         =   "未完成"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   13815
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   180
            Width           =   945
         End
         Begin VB.OptionButton ORD_NO_Y 
            BackColor       =   &H00E0E0E0&
            Caption         =   "完成"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   12750
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   180
            Value           =   -1  'True
            Width           =   900
         End
         Begin VB.TextBox txt_ord_kndname 
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
            Left            =   6930
            TabIndex        =   9
            Top             =   120
            Width           =   1635
         End
         Begin VB.TextBox txt_ord_knd 
            Alignment       =   2  'Center
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
            Left            =   6330
            MaxLength       =   1
            TabIndex        =   5
            Text            =   "A"
            Top             =   120
            Width           =   585
         End
         Begin VB.ComboBox TXT_ORD_ITEM 
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
            Left            =   11700
            TabIndex        =   4
            Top             =   120
            Width           =   660
         End
         Begin VB.TextBox TXT_ORD_NO 
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
            Left            =   10350
            MaxLength       =   11
            TabIndex        =   3
            Top             =   120
            Width           =   1350
         End
         Begin InDate.ULabel ULabel3 
            Height          =   315
            Left            =   150
            Top             =   120
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "用户交货期"
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
         Begin InDate.UDate UDA_DEL_TO 
            Height          =   315
            Left            =   3090
            TabIndex        =   6
            Tag             =   "交货期"
            Top             =   120
            Width           =   1440
            _ExtentX        =   2540
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
            MaxLength       =   10
         End
         Begin InDate.ULabel ULabel5 
            Height          =   315
            Left            =   9060
            Top             =   120
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "订单号"
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
         Begin InDate.UDate UDA_DEL_FR 
            Height          =   315
            Left            =   1440
            TabIndex        =   7
            Tag             =   "交货期"
            Top             =   120
            Width           =   1440
            _ExtentX        =   2540
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
            MaxLength       =   10
         End
         Begin InDate.ULabel ULabel6 
            Height          =   315
            Left            =   5040
            Top             =   120
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "订单种类"
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
            ForeColor       =   16711680
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "~"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   120
            Left            =   2940
            TabIndex        =   8
            Top             =   210
            Width           =   90
         End
      End
   End
End
Attribute VB_Name = "CKP1020C"
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
'-- Program ID        CKP1020C
'-- Document No       Q-00-0010(Specification)
'-- Designer          杨猛
'-- Coder             杨猛
'-- Date              2010.01.28
'-- Description
'-- 中板厂订单完成情况统计
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

'Dim pColumn2 As New Collection      'Spread Primary Key Collection
'Dim nColumn2 As New Collection      'Spread necessary Column Collection
'Dim mColumn2 As New Collection      'Spread Maxlength check Column Collection
'Dim iColumn2 As New Collection      'Spread Insert Column Collection
'Dim aColumn2 As New Collection      'Master -> Spread Column Collection
'Dim lColumn2 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim Sc1 As New Collection           'Spread Collection
'Dim sc2 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim ls_PChangeName                  'To Record P control Name

Const SPD_ORD_ITEM_NUM = 3
Const SPD_ORD_TOT_WGT = 4
Const SPD_ORD_REM_WGT_U = 5
Const SPD_ORD_REM_WGT_D = 6
Const SPD_ORD_SMS_WGT = 7
Const SPD_ORD_CCM_WGT = 8
Const SPD_ORD_MILL_WGT = 9
Const SPD_ORD_CUT_WGT = 10


Private Sub Form_Define()

   Dim i As Integer
   Dim iRow As Integer
   
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
      Call Gp_Ms_Collection(UDA_DEL_FR, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(UDA_DEL_TO, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_ord_knd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(TXT_ORD_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(TXT_ORD_ITEM, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(ORD_OK, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     
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
    
    For iRow = 3 To 10
        Call Gp_Sp_Collection(ss1, iRow, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
    Next iRow
   
   For i = 1 To ss1.MaxCols
      Call Gp_Sp_ColColor(ss1, i, , &H8000000B)
   Next i
   
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="CKP1020C.P_SREFER", Key:="P-R"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc"
    
    Sc1.Item("Spread").Col = 0
    Sc1.Item("Spread").ROW = 0
    Sc1.Item("Spread").Text = "◎"
       
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
   
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
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "K-System.INI", Me.Name)
    
    UDA_DEL_FR.Text = Mid(UDA_DEL_FR.Text, 1, 8) + "01"
    UDA_DEL_TO.Text = Format(DateAdd("m", 1, UDA_DEL_FR.Text), "YYYY-MM-DD")
    UDA_DEL_TO.Text = DateAdd("d", -1, UDA_DEL_TO.Text)
    
    txt_ord_knd.Text = "A"
    Call ORD_NO_Y_Click
    
    Screen.MousePointer = vbDefault
   
End Sub

Public Sub Form_Cls()
    
    Call Gf_Sp_Cls(Sc1)
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    UDA_DEL_FR.Text = Mid(UDA_DEL_FR.Text, 1, 8) + "01"
    UDA_DEL_TO.Text = Format(DateAdd("m", 1, UDA_DEL_FR.Text), "YYYY-MM-DD")
    UDA_DEL_TO.Text = DateAdd("d", -1, UDA_DEL_TO.Text)
    
    txt_ord_knd.Text = "A"
    Call ORD_NO_Y_Click
    
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        
End Sub

Public Sub Form_Ref()

    Dim iCount              As Integer
    Dim dOrd_item_Num       As Double
    Dim dOrd_Tot_Wgt        As Double
    Dim dOrd_Rem_Wgt_U      As Double
    Dim dOrd_Rem_Wgt_D      As Double
    Dim dOrd_Sms_Wgt        As Double
    Dim dOrd_Ccm_Wgt        As Double
    Dim dOrd_Mill_Wgt       As Double
    Dim dOrd_Cut_Wgt        As Double
        
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Nothing) Then
        Call Gp_Sp_EvenRowBackcolor(ss1)
        ss1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
    
    With ss1
    
        If .MaxRows < 1 Then
           Exit Sub
        End If
        .MaxRows = .MaxRows + 1
        
        For iCount = 1 To .MaxRows - 1
        
            .ROW = iCount
            '订单序列总数
            .Col = SPD_ORD_ITEM_NUM:              dOrd_item_Num = dOrd_item_Num + Val(.Text)
            '订单总重量
            .Col = SPD_ORD_TOT_WGT:               dOrd_Tot_Wgt = dOrd_Tot_Wgt + Val(.Text)
            '订单欠量（上限）
            .Col = SPD_ORD_REM_WGT_U:             dOrd_Rem_Wgt_U = dOrd_Rem_Wgt_U + Val(.Text)
            '订单欠量（下限）
            .Col = SPD_ORD_REM_WGT_D:             dOrd_Rem_Wgt_D = dOrd_Rem_Wgt_D + Val(.Text)
            '炼钢
            .Col = SPD_ORD_SMS_WGT:               dOrd_Sms_Wgt = dOrd_Sms_Wgt + Val(.Text)
            '连铸
            .Col = SPD_ORD_CCM_WGT:               dOrd_Ccm_Wgt = dOrd_Ccm_Wgt + Val(.Text)
            '轧钢等待
            .Col = SPD_ORD_MILL_WGT:              dOrd_Mill_Wgt = dOrd_Mill_Wgt + Val(.Text)
            '精整等待
            .Col = SPD_ORD_CUT_WGT:               dOrd_Cut_Wgt = dOrd_Cut_Wgt + Val(.Text)


        Next iCount
        
            .ROW = .MaxRows
            .Col = 1:                             .Text = "合计"
            
            '订单序列总数
            .Col = SPD_ORD_ITEM_NUM:              .Text = dOrd_item_Num
            '订单总重量
            .Col = SPD_ORD_TOT_WGT:               .Text = dOrd_Tot_Wgt
            '订单欠量（上限）
            .Col = SPD_ORD_REM_WGT_U:             .Text = dOrd_Rem_Wgt_U
            '订单欠量（下限）
            .Col = SPD_ORD_REM_WGT_D:             .Text = dOrd_Rem_Wgt_D
            '炼钢
            .Col = SPD_ORD_SMS_WGT:               .Text = dOrd_Sms_Wgt
            '连铸
            .Col = SPD_ORD_CCM_WGT:               .Text = dOrd_Ccm_Wgt
            '轧钢等待
            .Col = SPD_ORD_MILL_WGT:              .Text = dOrd_Mill_Wgt
            '精整等待
            .Col = SPD_ORD_CUT_WGT:               .Text = dOrd_Cut_Wgt
            
            Call Gp_Sp_BlockColor(ss1, 1, .MaxCols, .ROW, .ROW, , &HFFC0FF)
            
    End With
    
End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub Form_Exc()
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
End Sub

Public Sub Sp_Setting(ByVal sPname As Variant)

    Dim iRow As Integer

    With sPname

        .RowHeight(-1) = 13

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
        .ROW = 0: .Row2 = -1


        .BlockMode = True
        .FontBold = False
        .FontName = "SimSun"
        .FontSize = 10
        .BlockMode = False

        .Col = -1
        .ROW = 0
        .FontBold = True
    End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

   If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "K-System.INI", Me.Name)
    
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

Private Sub ORD_NO_N_Click()
    ORD_OK.Text = "N"
    ORD_NO_N.ForeColor = &HFF&
    ORD_NO_Y.ForeColor = &H808080
End Sub

Private Sub ORD_NO_Y_Click()
    ORD_OK.Text = "Y"
    ORD_NO_Y.ForeColor = &HFF&
    ORD_NO_N.ForeColor = &H808080
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)
    
    Sc1.Item("Spread").Col = 0
    Sc1.Item("Spread").ROW = 0
    Sc1.Item("Spread").Text = "◎"

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

Private Sub txt_ord_knd_Change()
    If Len(Trim(txt_ord_knd)) = txt_ord_knd.MaxLength Then
        txt_ord_kndname.Text = Gf_ComnNameFind(M_CN1, "B0009", Trim(txt_ord_knd.Text), 2)
    Else
        txt_ord_kndname.Text = ""
    End If
End Sub

Private Sub txt_ord_knd_DblClick()

    Call txt_ord_knd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_ord_knd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0009"
        DD.rControl.Add Item:=txt_ord_knd
        DD.rControl.Add Item:=txt_ord_kndname

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If
    
End Sub
