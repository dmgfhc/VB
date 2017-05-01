VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form AGT1100C 
   Caption         =   "在制品查询_AGT1100C"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "AGT1100C"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   12990
   ScaleWidth      =   21480
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
      Locked          =   -1  'True
      PaneTree        =   "AGT1100C.frx":0000
      Begin Threed.SSFrame SSFrame2 
         Height          =   540
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   953
         _Version        =   196609
         BackColor       =   14737632
         Begin VB.TextBox TXT_SP_CD 
            Height          =   270
            Left            =   14490
            TabIndex        =   7
            Top             =   180
            Visible         =   0   'False
            Width           =   255
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
            ItemData        =   "AGT1100C.frx":0052
            Left            =   7545
            List            =   "AGT1100C.frx":0054
            TabIndex        =   3
            Top             =   120
            Width           =   765
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
            ItemData        =   "AGT1100C.frx":0056
            Left            =   10470
            List            =   "AGT1100C.frx":0058
            TabIndex        =   2
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
            TabIndex        =   4
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
         Begin Threed.SSOption OPT_THK 
            Height          =   330
            Left            =   12930
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
            Caption         =   "厚度"
         End
         Begin Threed.SSOption OPT_STD 
            Height          =   330
            Left            =   11910
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
            Caption         =   "标准"
            Value           =   -1
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "~"
            Height          =   120
            Left            =   3540
            TabIndex        =   6
            Top             =   240
            Width           =   195
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   8505
         Left            =   0
         TabIndex        =   10
         Top             =   600
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   15002
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BackColor       =   14737632
         TabCaption(0)   =   "在制品"
         TabPicture(0)   =   "AGT1100C.frx":005A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "ss1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "吊下"
         TabPicture(1)   =   "AGT1100C.frx":0076
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "ss2"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin FPSpread.vaSpread ss1 
            Height          =   8205
            Left            =   0
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   300
            Width           =   15255
            _Version        =   393216
            _ExtentX        =   26908
            _ExtentY        =   14473
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
            MaxCols         =   14
            MaxRows         =   10
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "AGT1100C.frx":0092
         End
         Begin FPSpread.vaSpread ss2 
            Height          =   8205
            Left            =   -75000
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   300
            Width           =   15255
            _Version        =   393216
            _ExtentX        =   26908
            _ExtentY        =   14473
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
            MaxCols         =   12
            MaxRows         =   10
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "AGT1100C.frx":0975
         End
      End
   End
End
Attribute VB_Name = "AGT1100C"
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
'-- Program ID        AGT1100C
'-- Document No       Q-00-0010(Specification)
'-- Designer          YANGMENG
'-- Coder
'-- Date              2009.04.28
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

   Dim I As Integer
   Dim iRow As Integer
   
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Hsheet"
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
  Call Gp_Ms_Collection(txt_from_date, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_to_date, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(CBO_SHIFT, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(CBO_GROUP, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(TXT_SP_CD, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     
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
    
    For iRow = 1 To 14
        Call Gp_Sp_Collection(ss1, iRow, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
    Next iRow
   
   For I = 1 To ss1.MaxCols
      Call Gp_Sp_ColColor(ss1, I, , &H8000000B)
   Next I
   
    'Spread_Collection
    
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AGT1100C.P_SREFER", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    For iRow = 1 To 12
        Call Gp_Sp_Collection(ss2, iRow, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2, True)
    Next iRow
   
   For I = 1 To ss1.MaxCols
      Call Gp_Sp_ColColor(ss2, I, , &H8000000B)
   Next I
    
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AGT1100C.P_SREFER1", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").Row = 0
    sc1.Item("Spread").Text = "◎"
       
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

    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_ControlLock(Mc1("lControl"), True)
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    Call Gp_Sp_Setting(ss2)
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColGet(ss2, "G-System.INI", Me.Name)
    
    CBO_SHIFT.AddItem "1"
    CBO_SHIFT.AddItem "2"
    CBO_SHIFT.AddItem "3"
    
    CBO_GROUP.AddItem "A"
    CBO_GROUP.AddItem "B"
    CBO_GROUP.AddItem "C"
    CBO_GROUP.AddItem "D"
    
    OPT_STD.Value = True
    
    Screen.MousePointer = vbDefault
   
End Sub

Public Sub Form_Cls()
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Gp_Ms_Cls(Mc1("rControl"))
        
End Sub

Public Sub Form_Ref()

    Dim I, j    As Integer
    Dim GROUP   As String
    Dim wgt(14) As Variant
    
    If Trim(txt_from_date.RawData) = "" Or Trim(txt_to_date.RawData) = "" Then
       MsgBox "查询日期未输入!", vbCritical, "系统提示信息"
       Exit Sub
    End If
    
    Select Case SSTab1.Tab
           
           Case 0
     
                If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Nothing) Then
                    ss1.OperationMode = OperationModeNormal
                    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                End If
                
                With ss1
                     .Col = 1
                     .MaxRows = .MaxRows + 1
                     .Row = .MaxRows
                     .Text = "合计"
                     
                     For I = 1 To .MaxRows
                            For j = 2 To .MaxCols
                            
                                .Row = I
                                .Col = j
                                 
                                If I < .MaxRows Then
                                     If Val(.Text) = 0 Then
                                       .Text = ""
                                     Else
                                       wgt(j - 1) = wgt(j - 1) + Val(.Text)
                                     End If
                                Else
                                     .Text = wgt(j - 1)
                                End If
                            Next j
                     Next I
                End With
                        
           Case 1
           
                If Gf_Sp_Refer(M_CN1, sc2, Mc1, Mc1("nControl"), Nothing) Then
                    ss1.OperationMode = OperationModeNormal
                    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                End If
                
                With ss2
                
                     .Col = 1
                     .MaxRows = .MaxRows + 1
                     .Row = .MaxRows
                     .Text = "合计"
                     
                     For I = 1 To .MaxRows
                            For j = 2 To .MaxCols
                            
                                .Row = I
                                .Col = j
                                 
                                If I < .MaxRows Then
                                     If Val(.Text) = 0 Then
                                       .Text = ""
                                     Else
                                       wgt(j - 1) = wgt(j - 1) + Val(.Text)
                                     End If
                                Else
                                     .Text = wgt(j - 1)
                                End If
                            Next j
                     Next I
                End With
    
    End Select
    
End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub Form_Exc()

    If SSTab1.Tab = 0 Then
        Call Gp_Sp_Excel_AGT1100C0(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    Else
        Call Gp_Sp_Excel_AGT1100C1(Me, ss2, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    End If

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
    Call Gp_Sp_ColSet(ss2, "G-System.INI", Me.Name)
    
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
Private Sub OPT_THK_Click(Value As Integer)

    Dim iRow As Integer
    Dim sTemp As String

    If OPT_THK.Value = True Then
        OPT_THK.ForeColor = &HFF&
        OPT_STD.ForeColor = &H808080
        TXT_SP_CD = "T"
        ss1.Row = 0:        ss1.Col = 1:        ss1.Text = "厚度"
        ss2.Row = 0:        ss2.Col = 1:        ss2.Text = "班别"
        Call Gf_Sp_Cls(sc1)
        Call Gf_Sp_Cls(sc2)
    Else
        OPT_THK.ForeColor = &H808080
        TXT_SP_CD = "S"
    End If
    
End Sub

Private Sub OPT_STD_Click(Value As Integer)

    Dim iRow As Integer
    Dim sTemp As String

    If OPT_STD.Value = True Then
        OPT_STD.ForeColor = &HFF&
        OPT_THK.ForeColor = &H808080
        TXT_SP_CD = "S"
        ss1.Row = 0:        ss1.Col = 1:        ss1.Text = "标准号"
        ss2.Row = 0:        ss2.Col = 1:        ss2.Text = "标准号"
        Call Gf_Sp_Cls(sc1)
        Call Gf_Sp_Cls(sc2)
    Else
        OPT_STD.ForeColor = &H808080
        TXT_SP_CD = "T"
    End If
    
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)


'    lBlkcol1 = 0
'    lBlkcol2 = 0
'    lBlkrow1 = 1
'    lBlkrow2 = 0

    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2
    
End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
End Sub
'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_Excel
'   2.Name         : Spread --> Excel
'   3.Input  Value : Fm Form, sPname Variant, bLkcol1 Long, bLkcol2 Long, bLkrow1 Long, bLkrow2 Long
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread --> Excel
'---------------------------------------------------------------------------------------
Public Sub Gp_Sp_Excel_AGT1100C0(Fm As Form, sPname As Variant, bLkcol1 As Long, bLkcol2 As Long, bLkrow1 As Long, bLkrow2 As Long)

On Error GoTo Excel_Error

    Dim ret         As Boolean
    Dim xlApp       As Object
    Dim xlBpp       As Object
    Dim xlBook      As Object
    Dim xlSheet     As Object
    Dim ColIndex    As Integer
    Dim sExlRange   As String
    Dim sExlRange1  As String
    Dim iExlCol     As Integer
    
    With sPname
    
        If .MaxRows = 0 Then Exit Sub
        
        Clipboard.Clear
        
        .Col = 1: .Col2 = -1
        .Row = 1: .Row2 = -1
        
        Clipboard.SetText .Clip
        
        'Call Excel
        Set xlApp = CreateObject("Excel.Application")
        Set xlBook = xlApp.Workbooks.Add
        Set xlSheet = xlBook.Worksheets(1)
    
        xlApp.Visible = True
        
        xlSheet.Range("A1:A3").MergeCells = True:  xlSheet.Range("A1").Value = "标准号"
        xlSheet.Range("B1:I1").MergeCells = True:  xlSheet.Range("B1").Value = "在制品"
        xlSheet.Range("B2:B3").MergeCells = True:  xlSheet.Range("B2").Value = "等性能"
        xlSheet.Range("C2:C3").MergeCells = True:  xlSheet.Range("C2").Value = "探伤"
        xlSheet.Range("D2:D3").MergeCells = True:  xlSheet.Range("D2").Value = "堆冷"
        xlSheet.Range("E2:E3").MergeCells = True:  xlSheet.Range("E2").Value = "计划切割"
        xlSheet.Range("F2:H2").MergeCells = True:  xlSheet.Range("F2").Value = "热处理"
        xlSheet.Range("F3:F3").MergeCells = True:  xlSheet.Range("F3").Value = "正火"
        xlSheet.Range("G3:G3").MergeCells = True:  xlSheet.Range("G3").Value = "回火"
        xlSheet.Range("H3:H3").MergeCells = True:  xlSheet.Range("H3").Value = "淬火"
        xlSheet.Range("I2:I3").MergeCells = True:  xlSheet.Range("I2").Value = "协议"
        
        xlSheet.Range("J1:M1").MergeCells = True:  xlSheet.Range("J1").Value = "待处理"
        xlSheet.Range("J2:J3").MergeCells = True:  xlSheet.Range("J2").Value = "修磨"
        xlSheet.Range("K2:K3").MergeCells = True:  xlSheet.Range("K2").Value = "非计划毛边"
        xlSheet.Range("L2:L3").MergeCells = True:  xlSheet.Range("L2").Value = "矫直"
        xlSheet.Range("M2:M3").MergeCells = True:  xlSheet.Range("M2").Value = "性能挽救"
        xlSheet.Range("N1:N3").MergeCells = True:  xlSheet.Range("N1").Value = "未入库合计"
        
        xlSheet.Cells.NumberFormatLocal = "G/通用格式"
        xlSheet.Range("A4").Select
        xlSheet.Paste
        xlSheet.Cells.EntireColumn.AutoFit       'Column AutoFit
        
'        xlSheet.Range("A1:Q1").MergeCells = True
        
        sExlRange1 = ""
        For ColIndex = 1 To .MaxCols
            .Col = ColIndex
            .Row = 1
            
            iExlCol = ColIndex
            If IsNumeric(.Text) And Left(.Text, 1) = "0" And _
               (Len(.Text) = 8 Or Len(.Text) = 10 Or Len(.Text) = 12 Or Len(.Text) = 14) Then
                If ColIndex > 104 Then
                    sExlRange1 = "D" & sExlRange1
                    iExlCol = ColIndex - 104
                ElseIf ColIndex > 78 Then
                    sExlRange1 = "C" & sExlRange1
                    iExlCol = ColIndex - 78
                ElseIf ColIndex > 52 Then
                    sExlRange1 = "B" & sExlRange1
                    iExlCol = ColIndex - 52
                ElseIf ColIndex > 26 Then
                    sExlRange1 = "A"
                    iExlCol = ColIndex - 26
                End If
                
                sExlRange = sExlRange1 & Chr(iExlCol + 64) & "1:" & sExlRange1 & Chr(iExlCol + 64) & .MaxRows + 5
                If Len(.Text) = 8 Then
                    xlSheet.Range(sExlRange).NumberFormat = "00000000"
                ElseIf Len(.Text) = 10 Then
                    xlSheet.Range(sExlRange).NumberFormat = "0000000000"
                ElseIf Len(.Text) = 12 Then
                    xlSheet.Range(sExlRange).NumberFormat = "000000000000"
                ElseIf Len(.Text) = 14 Then
                    xlSheet.Range(sExlRange).NumberFormat = "00000000000000"
                End If
            End If
        Next
    
        Set xlSheet = Nothing
        Set xlBook = Nothing
        Set xlApp = Nothing
        
    End With
    
    Exit Sub
    
Excel_Error:

    Call Gp_MsgBoxDisplay("您的机器尚未安装Excel " & Error, "W")

End Sub
'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_Excel
'   2.Name         : Spread --> Excel
'   3.Input  Value : Fm Form, sPname Variant, bLkcol1 Long, bLkcol2 Long, bLkrow1 Long, bLkrow2 Long
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread --> Excel
'---------------------------------------------------------------------------------------
Public Sub Gp_Sp_Excel_AGT1100C1(Fm As Form, sPname As Variant, bLkcol1 As Long, bLkcol2 As Long, bLkrow1 As Long, bLkrow2 As Long)

On Error GoTo Excel_Error

    Dim ret         As Boolean
    Dim xlApp       As Object
    Dim xlBpp       As Object
    Dim xlBook      As Object
    Dim xlSheet     As Object
    Dim ColIndex    As Integer
    Dim sExlRange   As String
    Dim sExlRange1  As String
    Dim iExlCol     As Integer
    
    With sPname
    
        If .MaxRows = 0 Then Exit Sub
        
        Clipboard.Clear
        
        .Col = 1: .Col2 = -1
        .Row = 1: .Row2 = -1
        
        Clipboard.SetText .Clip
        
        'Call Excel
        Set xlApp = CreateObject("Excel.Application")
        Set xlBook = xlApp.Workbooks.Add
        Set xlSheet = xlBook.Worksheets(1)
    
        xlApp.Visible = True
        
        xlSheet.Range("A1:A2").MergeCells = True:  xlSheet.Range("A1").Value = "标准号"
        xlSheet.Range("B1:B2").MergeCells = True:  xlSheet.Range("B1").Value = "吊下总量"
        xlSheet.Range("C1:D1").MergeCells = True:  xlSheet.Range("C1").Value = "修磨"
        xlSheet.Range("C2:C2").MergeCells = True:  xlSheet.Range("C2").Value = "当日"
        xlSheet.Range("D2:D2").MergeCells = True:  xlSheet.Range("D2").Value = "累计"
        xlSheet.Range("E1:F1").MergeCells = True:  xlSheet.Range("E1").Value = "计划外毛边"
        xlSheet.Range("E2:E2").MergeCells = True:  xlSheet.Range("E2").Value = "当日"
        xlSheet.Range("F2:F2").MergeCells = True:  xlSheet.Range("F2").Value = "累计"
        xlSheet.Range("G1:H1").MergeCells = True:  xlSheet.Range("G1").Value = "计划内毛边"
        xlSheet.Range("G2:G2").MergeCells = True:  xlSheet.Range("G2").Value = "当日"
        xlSheet.Range("H2:H2").MergeCells = True:  xlSheet.Range("H2").Value = "累计"
        xlSheet.Range("I1:J1").MergeCells = True:  xlSheet.Range("I1").Value = "瓢曲"
        xlSheet.Range("I2:I2").MergeCells = True:  xlSheet.Range("I2").Value = "当日"
        xlSheet.Range("J2:J2").MergeCells = True:  xlSheet.Range("J2").Value = "累计"
        xlSheet.Range("K1:L1").MergeCells = True:  xlSheet.Range("K1").Value = "待热处理"
        xlSheet.Range("K2:K2").MergeCells = True:  xlSheet.Range("K2").Value = "当日"
        xlSheet.Range("L2:L2").MergeCells = True:  xlSheet.Range("L2").Value = "累计"
        
        xlSheet.Cells.NumberFormatLocal = "G/通用格式"
        xlSheet.Range("A3").Select
        xlSheet.Paste
        xlSheet.Cells.EntireColumn.AutoFit       'Column AutoFit
        
'        xlSheet.Range("A1:Q1").MergeCells = True
        
        sExlRange1 = ""
        For ColIndex = 1 To .MaxCols
            .Col = ColIndex
            .Row = 1
            
            iExlCol = ColIndex
            If IsNumeric(.Text) And Left(.Text, 1) = "0" And _
               (Len(.Text) = 8 Or Len(.Text) = 10 Or Len(.Text) = 12 Or Len(.Text) = 14) Then
                If ColIndex > 104 Then
                    sExlRange1 = "D" & sExlRange1
                    iExlCol = ColIndex - 104
                ElseIf ColIndex > 78 Then
                    sExlRange1 = "C" & sExlRange1
                    iExlCol = ColIndex - 78
                ElseIf ColIndex > 52 Then
                    sExlRange1 = "B" & sExlRange1
                    iExlCol = ColIndex - 52
                ElseIf ColIndex > 26 Then
                    sExlRange1 = "A"
                    iExlCol = ColIndex - 26
                End If
                
                sExlRange = sExlRange1 & Chr(iExlCol + 64) & "1:" & sExlRange1 & Chr(iExlCol + 64) & .MaxRows + 5
                If Len(.Text) = 8 Then
                    xlSheet.Range(sExlRange).NumberFormat = "00000000"
                ElseIf Len(.Text) = 10 Then
                    xlSheet.Range(sExlRange).NumberFormat = "0000000000"
                ElseIf Len(.Text) = 12 Then
                    xlSheet.Range(sExlRange).NumberFormat = "000000000000"
                ElseIf Len(.Text) = 14 Then
                    xlSheet.Range(sExlRange).NumberFormat = "00000000000000"
                End If
            End If
        Next
    
        Set xlSheet = Nothing
        Set xlBook = Nothing
        Set xlApp = Nothing
        
    End With
    
    Exit Sub
    
Excel_Error:

    Call Gp_MsgBoxDisplay("您的机器尚未安装Excel " & Error, "W")

End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)

If PreviousTab = 0 Then
    Call Gf_Sp_Cls(sc2)
    OPT_THK.Caption = "班别"
Else
    Call Gf_Sp_Cls(sc1)
    OPT_THK.Caption = "厚度"
End If
        
End Sub

