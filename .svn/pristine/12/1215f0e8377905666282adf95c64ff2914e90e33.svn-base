VERSION 5.00
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form AGT1050C 
   Caption         =   "非计划查询_AGT1050C"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17175
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   17175
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9105
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   16060
      _Version        =   196609
      SplitterBarWidth=   3
      BorderStyle     =   0
      PaneTree        =   "AGT1050C.frx":0000
      Begin FPSpread.vaSpread ss1 
         Height          =   8145
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   960
         Width           =   15255
         _Version        =   393216
         _ExtentX        =   26908
         _ExtentY        =   14367
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
         MaxCols         =   94
         MaxRows         =   3
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AGT1050C.frx":0052
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   900
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   1588
         _Version        =   196609
         BackColor       =   14737632
         Begin VB.TextBox TXT_ORG_ORD 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   13110
            MaxLength       =   2
            TabIndex        =   12
            Tag             =   "标准代码"
            Top             =   510
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox TXT_SP_CD 
            Height          =   270
            Left            =   14610
            TabIndex        =   11
            Top             =   180
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox chk_Org_Ord 
            BackColor       =   &H00E0E0E0&
            Caption         =   "原始订单号"
            Height          =   255
            Left            =   11790
            TabIndex        =   10
            Tag             =   "B"
            Top             =   540
            Width           =   1230
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
            ItemData        =   "AGT1050C.frx":2CDD
            Left            =   10470
            List            =   "AGT1050C.frx":2CDF
            TabIndex        =   4
            Top             =   120
            Width           =   765
         End
         Begin VB.ComboBox CBO_SHIFT 
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
            ItemData        =   "AGT1050C.frx":2CE1
            Left            =   7545
            List            =   "AGT1050C.frx":2CE3
            TabIndex        =   3
            Top             =   120
            Width           =   765
         End
         Begin InDate.ULabel ULabel11 
            Height          =   315
            Left            =   240
            Top             =   120
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   556
            Caption         =   "轧制时间"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.76
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin InDate.UDate txt_to_date 
            Height          =   315
            Left            =   3735
            TabIndex        =   5
            Top             =   120
            Width           =   1500
            _ExtentX        =   2646
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
         Begin InDate.UDate txt_from_date 
            Height          =   315
            Left            =   1950
            TabIndex        =   6
            Top             =   120
            Width           =   1500
            _ExtentX        =   2646
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
         Begin InDate.ULabel ULabel4 
            Height          =   315
            Left            =   6540
            Top             =   120
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            Caption         =   "班次"
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
         Begin InDate.ULabel ULabel5 
            Height          =   315
            Left            =   9465
            Top             =   120
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            Caption         =   "班别"
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
         Begin Threed.SSOption OPT_INSPSCRAP 
            Height          =   330
            Left            =   13050
            TabIndex        =   8
            Top             =   120
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   582
            _Version        =   196609
            Font3D          =   2
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
            Caption         =   "检验废"
         End
         Begin Threed.SSOption OPT_NOORD 
            Height          =   330
            Left            =   11790
            TabIndex        =   9
            Top             =   120
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   582
            _Version        =   196609
            Font3D          =   2
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
            Caption         =   "非计划"
            Value           =   -1
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "~"
            Height          =   120
            Left            =   3540
            TabIndex        =   7
            Top             =   240
            Width           =   195
         End
      End
   End
End
Attribute VB_Name = "AGT1050C"
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
'-- Program ID        AGT1050C
'-- Document No       Q-00-0010(Specification)
'-- Designer          GUOLI
'-- Coder
'-- Date              2008.02.13
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER      DATE         EDITOR       DESCRIPTION
'-- 1.01     2008.02.13   GUOLI
'-- 1.02     2012.07.17   LiQian       增加按原始订单查询非计划
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
Dim sc1 As New Collection           'Spread Collection
'Dim sc2 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim ls_PChangeName                  'To Record P control Name
 
Const SS1_STLGRD = 1   '轧制钢种
Const SS1_ORG_ORD = 2  '原始订单

Private Sub Form_Define()

   Dim i As Integer
   Dim iRow As Integer
   
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Hsheet"
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
  Call Gp_Ms_Collection(txt_from_date, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_to_date, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(CBO_SHIFT, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(CBO_GROUP, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(TXT_SP_CD, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(TXT_ORG_ORD, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) 'Add by LiQian at 2012-07-17 按原始订单号查询非计划
     
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
    
    For iRow = 3 To 93
        Call Gp_Sp_Collection(ss1, iRow, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
    Next iRow
   
   For i = 1 To ss1.MaxCols
      Call Gp_Sp_ColColor(ss1, i, , &H8000000B)
   Next i
   
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AGT1050C.P_SREFER", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").Row = 0
    sc1.Item("Spread").Text = "◎"
       
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
   
End Sub
'Add by LiQian at 2012-07-17 按原始订单查非计划
Private Sub chk_Org_Ord_Click()
   If chk_Org_Ord Then
        chk_Org_Ord.ForeColor = &HFF&
        TXT_ORG_ORD.Text = "Y"
        Call Form_Cls
        Call Gp_Sp_ColHidden(ss1, SS1_STLGRD, True)
        Call Gp_Sp_ColHidden(ss1, SS1_ORG_ORD, False)
    Else
        chk_Org_Ord.ForeColor = &H808080
        TXT_ORG_ORD.Text = ""
        Call Form_Cls
        Call Gp_Sp_ColHidden(ss1, SS1_STLGRD, False)
        Call Gp_Sp_ColHidden(ss1, SS1_ORG_ORD, True)
   End If
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
    
    Call Sp_Setting(Proc_Sc("Sc")("Spread"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "G-System.INI", Me.Name)
    
    Call Gp_Sp_ColHidden(ss1, SS1_ORG_ORD, True)
    chk_Org_Ord.Enabled = False
    
    CBO_SHIFT.AddItem "1"
    CBO_SHIFT.AddItem "2"
    CBO_SHIFT.AddItem "3"
    
    CBO_GROUP.AddItem "A"
    CBO_GROUP.AddItem "B"
    CBO_GROUP.AddItem "C"
    CBO_GROUP.AddItem "D"
    
    OPT_NOORD.Value = True
    
    Screen.MousePointer = vbDefault
   
End Sub

Public Sub Form_Cls()
    
    Call Gf_Sp_Cls(sc1)
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Gp_Ms_Cls(Mc1("rControl"))
        
End Sub

Public Sub Form_Ref()

    Dim i, j    As Integer
    Dim GROUP   As String
    Dim wgt(93) As Variant
    
    If Trim(txt_from_date.RawData) = "" Or Trim(txt_to_date.RawData) = "" Then
       MsgBox "查询日期未输入!", vbCritical, "系统提示信息"
       Exit Sub
    End If
        
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Nothing) Then
        With ss1
             .Col = 0
             .MaxRows = .MaxRows + 1
             .Row = .MaxRows
             .Text = "合计"
             
             For i = 1 To .MaxRows
                    For j = 3 To .MaxCols    '合计由2列->3列
                    
                        .Row = i
                        .Col = j
                         
                        If i < .MaxRows Then
                             If Val(.Text) = 0 Then
                               .Text = ""
                             Else
                               wgt(j - 1) = wgt(j - 1) + Val(.Text)
                             End If
                        Else
                             .Text = wgt(j - 1)
                        End If
                    Next j
             Next i
        End With
    End If
    
    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    
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

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

   If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "G-System.INI", Me.Name)
    
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

Private Sub OPT_INSPSCRAP_Click(Value As Integer)

    Dim iRow As Integer
    Dim sTemp As String

    If OPT_INSPSCRAP.Value = True Then
        OPT_INSPSCRAP.ForeColor = &HFF&
        OPT_NOORD.ForeColor = &H808080
        TXT_SP_CD = "I"
        ss1.Row = 0
        ss1.Col = 5                 '性能废由4列->5列
        ss1.Text = "性能废(t)"
        Call Gf_Sp_Cls(sc1)
        chk_Org_Ord.Value = False
        chk_Org_Ord.Enabled = False
        Call Gp_Sp_ColHidden(ss1, SS1_STLGRD, False)
        Call Gp_Sp_ColHidden(ss1, SS1_ORG_ORD, True)
    Else
        OPT_INSPSCRAP.ForeColor = &H808080
        TXT_SP_CD = "N"
    End If
    
End Sub

Private Sub OPT_NOORD_Click(Value As Integer)

    Dim iRow As Integer
    Dim sTemp As String

    If OPT_NOORD.Value = True Then
        OPT_NOORD.ForeColor = &HFF&
        OPT_INSPSCRAP.ForeColor = &H808080
        TXT_SP_CD = "N"
        ss1.Row = 0
        ss1.Col = 5               '非计划由4列->5列
        ss1.Text = "非计划总量(t)"
        Call Gf_Sp_Cls(sc1)
        chk_Org_Ord.Enabled = True
    Else
        OPT_NOORD.ForeColor = &H808080
        TXT_SP_CD = "I"
        chk_Org_Ord.Enabled = False
    End If
    
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").Row = 0
    sc1.Item("Spread").Text = "◎"

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
'    If Row > 0 Then
'        Set Active_Spread = Me.ss1
'        PopupMenu MDIMain.PopUp_Spread
'    End If
End Sub

