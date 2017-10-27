VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form ACF0100C 
   Caption         =   "热装热送统计报表_ACF0100C"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   WindowState     =   2  'Maximized
   Begin InDate.ULabel ULabel3 
      DragMode        =   1  'Automatic
      Height          =   315
      Left            =   120
      Top             =   120
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "订单月份"
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
   Begin InDate.UDate prod_date_from 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Tag             =   "开始日期"
      Top             =   120
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Text            =   "____-__"
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
      Mask            =   "%%%%-%%"
      MaxLength       =   7
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9855
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   20295
      _ExtentX        =   35798
      _ExtentY        =   17383
      _Version        =   196609
      PaneTree        =   "ACF0100C.frx":0000
      Begin Threed.SSPanel SSPanel2 
         Height          =   9795
         Left            =   8100
         TabIndex        =   2
         Top             =   30
         Width           =   12165
         _ExtentX        =   21458
         _ExtentY        =   17277
         _Version        =   196609
         Caption         =   "SSPanel2"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin FPSpread.vaSpread ss2 
            Height          =   10215
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   11595
            _Version        =   393216
            _ExtentX        =   20452
            _ExtentY        =   18018
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
            MaxCols         =   6
            MaxRows         =   13
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "ACF0100C.frx":0052
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   9795
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   7980
         _ExtentX        =   14076
         _ExtentY        =   17277
         _Version        =   196609
         Caption         =   "SSPanel1"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin FPSpread.vaSpread ss1 
            Height          =   8490
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   7875
            _Version        =   393216
            _ExtentX        =   13891
            _ExtentY        =   14975
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
            MaxCols         =   4
            MaxRows         =   36
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "ACF0100C.frx":05DC
         End
      End
   End
End
Attribute VB_Name = "ACF0100C"
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
'-- Program ID        ACF0100C
'-- Document No       Q-00-0010(Specification)
'-- Designer          WL
'-- Coder             WL
'-- Date              2017.8.22
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


'Const SS2_PLT = 1




Dim sWgtLenFlag As String
Dim sQuery  As String

Private Sub Form_Define()

 Dim iRow As Integer
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

     Call Gp_Ms_Collection(prod_date_from, "p", "n", "", " ", "r", " ", "", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
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
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)

    

    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="ACF0100C.P_REFER1", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc1, Key:="Sc1"
    
    
'    sc1.Item("Spread").Col = 0
'    sc1.Item("Spread").Row = 0
'    sc1.Item("Spread").Text = "◎"


    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     
     For iRow = 5 To ss2.MaxCols
        Call Gp_Sp_Collection(ss2, iRow, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2, True)
     Next iRow
    

    
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="ACF0100C.P_REFER2", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc2, Key:="Sc2"
    
    'Call Gp_Sp_ColHidden(ss2, 13, True)
    
    
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    
    
    
'    Sc3.Item("Spread").Col = 0
'    Sc3.Item("Spread").Row = 0
'    Sc3.Item("Spread").Text = "◎"
    
        
End Sub



Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)

End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    Call Form_Define
        
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Sp_Setting(Proc_Sc("Sc1")("Spread"))
    Call Sp_Setting(Proc_Sc("Sc2")("Spread"))

    Call Gp_Spl_SizeGet(SSSplitter1, "K-System.INI", Me.Name, "H")
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc1")("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc2")("Spread"), "K-System.INI", Me.Name)


     prod_date_from.Text = Mid(Date, 1, 7)
    
    Screen.MousePointer = vbDefault

End Sub
    

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc1")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Spl_SizeSet(SSSplitter1, "C-System.INI", Me.Name)
    
    Call Gp_Sp_ColSet(sc1.Item("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "C-System.INI", Me.Name)

    
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

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

Public Sub Form_Cls()
    
    ss1.ClearRange 1, 1, ss1.MaxCols, ss1.MaxRows, True
    ss2.ClearRange 1, 1, ss2.MaxCols, ss2.MaxRows, True
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
'    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    
    With MDIMain.MenuTool
        .Buttons(5).Enabled = False                 'Delete
        .Buttons(6).Enabled = False                 'Separator
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = False                 'Row Delete
        .Buttons(9).Enabled = False                 'Row Cancel
        .Buttons(10).Enabled = False                'Separator
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
    End With
    
End Sub

Public Sub Form_Ref()

    Dim i As Integer
    Dim ORD_WGT  As Double
    Dim HOT_WGT  As Double
    Dim ACT_WGT  As Double
    Dim ACT_RATE As Double
    Dim HOT_RATE As Double
    Dim ORD_WGT_C1 As Double
    Dim ORD_WGT_C2 As Double
    Dim ORD_WGT_C3 As Double
    Dim HOT_WGT_C3_C As Double
    Dim HOT_WGT_C3_H As Double
    Dim HOT_WGT_C2_C As Double
    Dim HOT_WGT_C2_H As Double
    Dim HOT_WGT_C1_C As Double
    Dim HOT_WGT_C1_H As Double
    Dim ACT_WGT_C3 As Double
    Dim ACT_WGT_C2 As Double
    Dim ACT_WGT_C1 As Double
    Dim SLAB_WGT_C2 As Double
    Dim SLAB_WGT_C1 As Double
    Dim SLAB_WGT_C3 As Double
    

      On Error Resume Next

    Call Form_Cls
    
    ss1.ReDraw = False
    ss2.ReDraw = False
   
'    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc1"), Mc1, Mc1("nControl")) Then
    If Sp_Display(M_CN1, Proc_Sc("Sc1")("Spread"), Gf_Ms_MakeQuery(Proc_Sc("Sc1").Item("P-R"), "R", Mc1("pControl"))) Then
        Call Sp_Display(M_CN1, Proc_Sc("Sc2")("Spread"), Gf_Ms_MakeQuery(Proc_Sc("Sc2").Item("P-R"), "R", Mc1("pControl")))
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
    
    With MDIMain.MenuTool
        .Buttons(5).Enabled = False                 'Delete
        .Buttons(6).Enabled = False                 'Separator
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = False                 'Row Delete
        .Buttons(9).Enabled = False                 'Row Cancel
        .Buttons(10).Enabled = False                'Separator
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
    End With
    ss1.ReDraw = True
    ss2.ReDraw = True
        
    
    With ss1
         For i = 1 To .MaxCols
         .Col = i
         .Row = 5
         .BackColor = &HC0C0FF
         .Row = 12
         .BackColor = &HC0C0FF
         .Row = 21
         .BackColor = &HC0C0FF
         .Row = 25
         .BackColor = &HC0C0FF
         .Row = 29
         .BackColor = &HC0C0FF
         .Row = 35
         .BackColor = &HC0C0FF
         .Row = 36
         .BackColor = &HC0C0FF
         Next
         '订单量
         .Col = 2
         .Row = 4:    ORD_WGT_C3 = Val(.Text)
         .Row = 5:   .Text = ORD_WGT_C3
         .Row = 12:  .Text = ORD_WGT_C3
         .Row = 20:   ORD_WGT_C1 = Val(.Text)
         .Row = 21:  .Text = ORD_WGT_C1
         .Row = 25:  .Text = ORD_WGT_C1
         .Row = 28:   ORD_WGT_C2 = Val(.Text)
         .Row = 29:  .Text = ORD_WGT_C2
         .Row = 35:  .Text = ORD_WGT_C2
         .Row = 36:  .Text = ORD_WGT_C1 + ORD_WGT_C2 + ORD_WGT_C3
         .BackColor = &HC0C0FF
         
         '可热装量
         .Col = 3
         For i = 1 To 4
         .Row = i:    HOT_WGT_C3_H = HOT_WGT_C3_H + Val(.Text)
         Next
         For i = 6 To 11
         .Row = i:    HOT_WGT_C3_C = HOT_WGT_C3_C + Val(.Text)
         Next
         For i = 13 To 20
         .Row = i:    HOT_WGT_C1_H = HOT_WGT_C1_H + Val(.Text)
         Next
         For i = 22 To 24
         .Row = i:    HOT_WGT_C1_C = HOT_WGT_C1_C + Val(.Text)
         Next
         For i = 26 To 28
         .Row = i:    HOT_WGT_C2_H = HOT_WGT_C2_H + Val(.Text)
         Next
         For i = 30 To 34
         .Row = i:    HOT_WGT_C2_C = HOT_WGT_C2_C + Val(.Text)
         Next
         .Row = 5:     .Text = HOT_WGT_C3_H
         .Row = 12:    .Text = HOT_WGT_C3_C
         .Row = 21:    .Text = HOT_WGT_C1_H
         .Row = 25:    .Text = HOT_WGT_C1_C
         .Row = 29:    .Text = HOT_WGT_C2_H
         .Row = 35:    .Text = HOT_WGT_C2_C
         .Row = 36:    .Text = HOT_WGT_C3_H + HOT_WGT_C3_C + HOT_WGT_C1_H + HOT_WGT_C2_H + HOT_WGT_C1_C + HOT_WGT_C2_C
    End With
        
    With ss1
         For i = 1 To .MaxRows
         .Row = i
         .Col = 3:   HOT_WGT = Val(.Text)  '可热装量
         .Col = 2:   ORD_WGT = Val(.Text)  '订单量量
         '可热装率计算
         If ORD_WGT = 0 Then
            HOT_RATE = 0
         Else
            HOT_RATE = HOT_WGT / ORD_WGT
         End If
         
         .Col = 4:   .Text = Round(HOT_RATE, 2) '可热装率
         Next
    End With
    
    With ss2
         .Col = 3
         For i = 1 To 3
         .Row = i: ACT_WGT_C3 = ACT_WGT_C3 + Val(.Text)
         Next
         For i = 5 To 7
         .Row = i: ACT_WGT_C1 = ACT_WGT_C1 + Val(.Text)
         Next
         For i = 9 To 11
         .Row = i: ACT_WGT_C2 = ACT_WGT_C2 + Val(.Text)
         Next
         .Row = 4:  .Text = ACT_WGT_C3
         .Row = 8:  .Text = ACT_WGT_C1
         .Row = 12: .Text = ACT_WGT_C2
         .Row = 13: .Text = ACT_WGT_C3 + ACT_WGT_C1 + ACT_WGT_C2
         
         .Col = 2
         .Row = 1:   SLAB_WGT_C3 = Val(.Text)
         .Row = 5:   SLAB_WGT_C1 = Val(.Text)
         .Row = 9:   SLAB_WGT_C2 = Val(.Text)
         .Row = 4:  .Text = SLAB_WGT_C3
         .Row = 8:  .Text = SLAB_WGT_C1
         .Row = 12: .Text = SLAB_WGT_C2
         .Row = 13: .Text = SLAB_WGT_C3 + SLAB_WGT_C2 + SLAB_WGT_C1
         
    End With
    
    With ss2
         For i = 1 To .MaxRows
         .Row = i
         .Col = 5:   HOT_WGT = Val(.Text)  '可热装量
         .Col = 3:   ACT_WGT = Val(.Text)  '实际热装量
         .Col = 2:   ORD_WGT = Val(.Text)  '坯料量量
         '可热装率计算
         If ORD_WGT = 0 Then
            HOT_RATE = 0
         Else
            HOT_RATE = HOT_WGT / ORD_WGT
         End If
         
         '实际热装率计算
         If ORD_WGT = 0 Then
            ACT_RATE = 0
         Else
            ACT_RATE = ACT_WGT / ORD_WGT
         End If
         .Col = 4:   .Text = Round(ACT_RATE, 2) '实际热装率
         .Col = 6:   .Text = Round(HOT_RATE, 2) '可热装率
         Next
     End With
                  
        
End Sub


'Public Sub Spread_ColumnsSort()
'
'    Spread_ColSort.Show 1
'
'End Sub

Public Sub Form_Exc()




   Dim i               As Integer
    Dim j               As Integer
    Dim xlApp           As Object
    Dim xlSheet         As Object
    Dim sRow            As String
    Dim sDate           As String
    Dim sShift          As String
    
    Dim sPage_Num       As Integer
    Dim sPage_X         As Integer
    Dim sPage           As Double
    Dim sLastPage       As Double
    Dim sRow1           As Integer
    Dim sRow2           As Integer
    
    Dim xl_1            As String
    Dim xl_2            As String
    Dim xl_3            As String
    Dim xl_4            As String
    Dim xl_5            As String
    Dim xl_6            As String
    Dim xl_7            As String
    'Dim xl_H            As String
    'Dim xl_I            As String
'    Dim xl_J            As String
    
    Dim xl_clr_body     As String
    Dim xl_clr_sum      As String
    Dim xl_clr_spc      As String
    
    Dim Xl_Cnt          As String
    Dim Xl_Wgt          As String
    Dim Xl_Wgt_Val      As String
    Dim Xl_Ust          As String
    
    If ss1.MaxRows < 1 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
     
    On Error Resume Next
    
    Set xlApp = GetObject(, "Excel.Application")
    If ERR.Number <> 0 Then
        Set xlApp = CreateObject("Excel.Application")
    End If
    
    ERR.Clear

    xlApp.Workbooks.Open (App.Path & "\ACF0100C.xls")
    
    Set xlSheet = xlApp.Worksheets("Sheet1")
    xlApp.Sheets("Sheet1").Select
    
    sDate = prod_date_from.Text
    
        xlApp.Range("B4").Value = Left(sDate, 4) + "年" + Mid(sDate, 6, 2) + "月" + Mid(sDate, 9, 2) + "日"
    
    

    xlApp.Range("D43").Value = Now
    
    xlApp.Range("O43").Value = sUserName
        
xlApp.Application.Visible = True
    

        xl_1 = "B6:E41"
        xl_2 = "I6:N18"
        
        Clipboard.Clear
        ss1.SetSelection 1, 1, ss1.MaxCols, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range(xl_1).Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        ss1.ClearSelection
        Sleep 100
        
        Clipboard.Clear
        ss2.SetSelection 1, 1, ss2.MaxCols, ss2.MaxRows
        ss2.ClipboardCopy
        xlApp.Range(xl_2).Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        ss2.ClearSelection
        Sleep 100
        
    
        

       
    Screen.MousePointer = vbDefault
    
    'xlApp.Application.Visible = True
    
    Set xlSheet = Nothing
    Set xlApp = Nothing
        
    Exit Sub

ErrHandle:
    MsgBox Error
    Set xlSheet = Nothing
    Set xlApp = Nothing
    Screen.MousePointer = vbDefault


End Sub

Public Sub Sp_Setting(ByVal sPname As Variant)

    Dim iRow As Integer

    With sPname
        .RowHeight(-1) = 13
        
        If .ColHeaderRows > 1 Then
            .RowHeight(SpreadHeader + (.ColHeaderRows - 2)) = 13
            .RowHeight(SpreadHeader + (.ColHeaderRows - 1)) = 13
        Else
            .RowHeight(0) = 24
        End If
        
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

Public Function Sp_Display(Conn As ADODB.Connection, sPname As Variant, sQuery As String) As Boolean

    On Error Resume Next

    Dim iCount          As Integer
    Dim iRowCount       As Long
    Dim iColcount       As Long
    Dim AdoRs           As ADODB.Recordset
    Dim ArrayRecords    As Variant

    Sp_Display = True

    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then Sp_Display = False: Exit Function
    End If

    Set AdoRs = New ADODB.Recordset

    With sPname

        .ReDraw = False
        iCount = 0

'        .ClearRange 1, 1, .MaxCols, .MaxRows, True

        Screen.MousePointer = vbHourglass

        'Ado Execute
        AdoRs.Open sQuery, Conn, adOpenKeyset

        If AdoRs.BOF Or AdoRs.EOF Then

            .ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
            Sp_Display = False
            Call Gp_MsgBoxDisplay("无相关记录", "I")
            Call Form_Cls
            Screen.MousePointer = vbDefault
            Exit Function

        End If
        
        ArrayRecords = AdoRs.GetRows
        
        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 2) <> 0 Then

            For iRowCount = 0 To .MaxRows - 1
'                Select Case Trim(ArrayRecords(0, iRowCount))
'                    Case "A0"
'                        .Row = 1
'                    Case "A1"
'                        .Row = 2
'                    Case "B0"
'                        .Row = 3
'                    Case "B1"
'                        .Row = 4
'                    Case "C0"
'                        .Row = 5
'                    Case "C1"
'                        .Row = 6
'                    Case "D0"
'                        .Row = 7
'                    Case "D1"
'                        .Row = 8
'                    Case "T0"
'                        .Row = 9
'                    Case "T1"
'                        .Row = 10
'                End Select
                 .Row = Val(Trim(ArrayRecords(0, iRowCount)))
            
'            .ROW = iRowCount + 1

                For iColcount = 1 To .MaxCols
    
                    .Col = iColcount
    
                    If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Then
                        .Text = ""
                    Else
                        .Text = Trim(ArrayRecords(iColcount, iRowCount))
                    End If

                Next iColcount

            Next iRowCount

        End If

        .ReDraw = True
        Screen.MousePointer = vbDefault

    End With

End Function
'
Public Sub Form_Exit()
    Unload Me
End Sub


