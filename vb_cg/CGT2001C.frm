VERSION 5.00
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form CGT2001C 
   Caption         =   "堆冷时间统计报表_CGT2001C"
   ClientHeight    =   10875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   12990
   ScaleWidth      =   21480
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_cool 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   6270
      MaxLength       =   5
      TabIndex        =   4
      Tag             =   "时间范围3"
      Top             =   540
      Width           =   1485
   End
   Begin VB.TextBox txt_cool 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   4537
      MaxLength       =   5
      TabIndex        =   3
      Tag             =   "时间范围2"
      Top             =   540
      Width           =   1485
   End
   Begin VB.TextBox txt_cool 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   2805
      MaxLength       =   5
      TabIndex        =   2
      Tag             =   "时间范围1"
      Top             =   540
      Width           =   1485
   End
   Begin VB.TextBox txt_groupby 
      Height          =   330
      Left            =   6870
      TabIndex        =   13
      Top             =   90
      Visible         =   0   'False
      Width           =   885
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   120
      Top             =   540
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   556
      Caption         =   "堆冷时间范围(小时)"
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
   Begin FPSpread.vaSpread ss1 
      Height          =   8205
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   15015
      _Version        =   393216
      _ExtentX        =   26485
      _ExtentY        =   14473
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
      MaxCols         =   0
      MaxRows         =   0
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "CGT2001C.frx":0000
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   735
      Left            =   7920
      TabIndex        =   11
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   1296
      _Version        =   196609
      ForeColor       =   16711680
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "汇总字段"
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "板坯长度"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   4431
         TabIndex        =   8
         Tag             =   ",B.DSC_DATE"
         Top             =   300
         Width           =   1080
      End
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "板坯宽度"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   3104
         TabIndex        =   7
         Tag             =   ",B.DSC_DATE"
         Top             =   300
         Width           =   1080
      End
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "板坯厚度"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1777
         TabIndex        =   6
         Tag             =   ",B.DSC_DATE"
         Top             =   300
         Width           =   1080
      End
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "订单标准"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   5760
         TabIndex        =   9
         Tag             =   ",SUBSTR(B.BED_PILE_DATE,1,8)"
         Top             =   300
         Width           =   1080
      End
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "板坯钢种"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   450
         TabIndex        =   5
         Tag             =   ",B.PROD_DATE"
         Top             =   300
         Width           =   1080
      End
   End
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   120
      Top             =   120
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   556
      Caption         =   "装  炉  时  间"
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
   Begin InDate.UDate ud_ch_date1 
      Height          =   315
      Left            =   2805
      TabIndex        =   0
      Tag             =   "起始日期"
      Top             =   120
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
   Begin InDate.UDate ud_ch_date2 
      Height          =   315
      Left            =   4537
      TabIndex        =   1
      Tag             =   "终止日期"
      Top             =   120
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
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "～"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   120
      Left            =   4320
      TabIndex        =   12
      Top             =   210
      Width           =   195
   End
End
Attribute VB_Name = "CGT2001C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       Nisco Production Management System
'-- Sub_System Name   Mill System
'-- Program Name      堆冷时间统计报表
'-- Program ID        CGT2001C
'-- Document No       Q-00-0010(Specification)
'-- Designer          GUOLI
'-- Date              2010.10.26
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

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
    
    Dim sQuery As String
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
    Call Gp_Ms_Collection(ud_ch_date1, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(ud_ch_date2, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_cool(1), "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_cool(2), "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_cool(3), "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    'Duplicate Count
    iDupCnt = 1
    
    'Sum Column Count
    iSumCnt = 2
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
        
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    Call Menu_Setting

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
    Call Gp_Ms_NeceColor(Mc1("nControl"))

    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    Call Sp_Setting
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "CG-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "CG-System.INI", Me.Name)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Form_Cls()
Dim i As Integer
    ss1.MaxCols = 0
    ss1.MaxRows = 0
    txt_groupby.Text = ""
    For i = 1 To 3
        chk_Cond(i).Value = UNCHECKED
    Next i
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Gp_Ms_ControlLock(Mc1("lControl"), False)

End Sub

Public Sub Form_Ref()

    Dim sMesg As String
    Dim i As Integer
    
    ss1.MaxRows = 0
    ss1.MaxCols = 0
    txt_groupby.Text = ""
    
    For i = 1 To 5
        If chk_Cond(i).Value = CHECKED Then
            txt_groupby.Text = txt_groupby.Text + Str(i)
        End If
    Next i
    
    sMesg = Gf_Ms_NeceCheck(nControl)
    If sMesg = "OK" Then
        If Sp_Header_Refer() Then
            If Sp_Data_Refer(ss1) Then
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            End If
        End If
    Else
        sMesg = sMesg + " 必须输入 ..."
        Call Gp_MsgBoxDisplay(sMesg)
    End If
    
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

Private Sub ss1_DblClick(ByVal Col As Long, ByVal ROW As Long)

Dim i As Integer
Dim tmp As Integer
Dim cur_row As Integer

    Unload CGT2001C_POP
    Load CGT2001C_POP
    CGT2001C_POP.Show
    CGT2001C_POP.txt_date1 = ud_ch_date1.RawData
    CGT2001C_POP.txt_date2 = ud_ch_date2.RawData
    cur_row = ss1.ActiveRow
    ss1.ROW = 0
    For i = 1 To 6
        ss1.Col = i
        
        If ss1.Text = chk_Cond(1).Caption Then
           ss1.ROW = cur_row
           CGT2001C_POP.txt_stlgrd = ss1.Text
        End If
        
        If ss1.Text = chk_Cond(2).Caption Then
           ss1.ROW = cur_row
           CGT2001C_POP.TXT_THK = ss1.Text
        End If
        
        If ss1.Text = chk_Cond(3).Caption Then
           ss1.ROW = cur_row
           CGT2001C_POP.TXT_WID = ss1.Text
        End If
        
        If ss1.Text = chk_Cond(4).Caption Then
           ss1.ROW = cur_row
           CGT2001C_POP.TXT_LEN = ss1.Text
        End If
        
        If ss1.Text = chk_Cond(5).Caption Then
           ss1.ROW = cur_row
           CGT2001C_POP.TXT_SPEC = ss1.Text
        End If
    Next i
    
    ss1.ROW = SpreadHeader
    ss1.Col = ss1.ActiveCol
    tmp = InStr(ss1.Text, " ")
    
    If InStr(ss1.Text, ">=") > 0 Then
       If tmp = 0 Then
          CGT2001C_POP.txt_cool1.Text = Mid(ss1.Text, 3)
       Else
          CGT2001C_POP.txt_cool1.Text = Mid(ss1.Text, 3, tmp - 3)
       End If
    Else
       CGT2001C_POP.txt_cool1.Text = 0
    End If
    
    If InStr(ss1.Text, "<") > 0 Then
       If tmp = 0 Then
          CGT2001C_POP.txt_cool2.Text = Mid(ss1.Text, 2)
       Else
          CGT2001C_POP.txt_cool2.Text = Mid(ss1.Text, tmp + 2)
       End If
    Else
       CGT2001C_POP.txt_cool2.Text = "99999"
    End If
    
    Call CGT2001C_POP.Form_Ref
    
End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal ROW As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If ROW > 0 Then
        Set Active_Spread = Me.ss1
        MDIMain.Mnu_Sorting.Visible = False
        MDIMain.Line1.Visible = False
        
        PopupMenu MDIMain.PopUp_Spread
        
        MDIMain.Mnu_Sorting.Visible = True
        MDIMain.Line1.Visible = True
    End If

End Sub

Public Sub Sp_Setting()

    With ss1

        .ColHeaderRows = 3
        .RowHeaderCols = 2
        
        .Col = -1
        .ROW = SpreadHeader + 1
        .FontBold = True
        
        .RowHeight(SpreadHeader) = 15
        .RowHeight(SpreadHeader + 1) = 15
        
        .ROW = SpreadHeader + 2
        .RowHidden = True
        
        .ColWidth(0) = 6
        
        .Col = 0
        .ColHidden = True
        
        .ColWidth(SpreadHeader + 1) = 10
        
        .Col = 0: .Col2 = -1
        .ROW = 0: .Row2 = 0
        
        .BlockMode = True
        .RowMerge = MergeAlways
        .ColMerge = MergeAlways
        .BlockMode = False
        
        .Col = SpreadHeader + 1: .Col2 = -1
        .ROW = 0: .Row2 = SpreadHeader + 1
        .BlockMode = True
        .RowMerge = MergeAlways
        .ColMerge = MergeAlways
        .BlockMode = False
        
        .ROW = SpreadHeader
        .Col = SpreadHeader + 1
        .Text = " "
        .ROW = SpreadHeader + 1
        .Col = SpreadHeader + 1
        .Text = " "
        
    End With

End Sub

Public Sub Menu_Setting()

    MDIMain.MenuTool.Buttons(5).Enabled = False    'Delete
    MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Inssert
    MDIMain.MenuTool.Buttons(8).Enabled = False    'Row delete
    MDIMain.MenuTool.Buttons(9).Enabled = False    'Row cancel
    MDIMain.MenuTool.Buttons(11).Enabled = False   'Row cancel
    MDIMain.MenuTool.Buttons(12).Enabled = False   'Row cancel
    
End Sub

Public Function Sp_Header_Refer() As Boolean

On Error GoTo SpreadDisplay_Error

    Dim i As Integer
    Dim j As Integer
    Dim cur_row As Integer
    
    Sp_Header_Refer = True
    
    With ss1
    
        For i = 1 To 3
            If txt_cool(i).Text <> "" Then
               If i = 1 Then
                    .MaxCols = .MaxCols + 2
                    .ROW = SpreadHeader
                    .Col = .MaxCols - 1
                    .Text = "<" & txt_cool(i).Text
                    .ROW = SpreadHeader
                    .Col = .MaxCols
                    .Text = "<" & txt_cool(i).Text
                    
                    .ROW = SpreadHeader + 1
                    .Col = .MaxCols - 1
                    .Text = "数量"
                    .ROW = SpreadHeader + 1
                    .Col = .MaxCols
                    .Text = "重量"
               Else
                    .MaxCols = .MaxCols + 2
                    .ROW = SpreadHeader
                    .Col = .MaxCols - 1
                    .Text = ">=" & txt_cool(i - 1).Text & " <" & txt_cool(i).Text
                    .ROW = SpreadHeader
                    .Col = .MaxCols
                    .Text = ">=" & txt_cool(i - 1).Text & " <" & txt_cool(i).Text
                    
                    .ROW = SpreadHeader + 1
                    .Col = .MaxCols - 1
                    .Text = "数量"
                    .ROW = SpreadHeader + 1
                    .Col = .MaxCols
                    .Text = "重量"
               End If
            Else
               Exit For
            End If
        Next i
        .MaxCols = .MaxCols + 2
        .ROW = SpreadHeader
        .Col = .MaxCols - 1
        .Text = ">=" & txt_cool(i - 1).Text
        .ROW = SpreadHeader
        .Col = .MaxCols
        .Text = ">=" & txt_cool(i - 1).Text
        
        .ROW = SpreadHeader + 1
        .Col = .MaxCols - 1
        .Text = "数量"
        .ROW = SpreadHeader + 1
        .Col = .MaxCols
        .Text = "重量"
               
        cur_row = .MaxCols
           
        For j = 1 To 5
            If chk_Cond(j).Value = CHECKED Then
               If j = 1 Then
                    .MaxCols = .MaxCols + 2
                    .InsertCols .MaxCols - cur_row - 1, 2
                    .Col = .MaxCols - cur_row - 1
                    .ROW = SpreadHeader
                    .Text = chk_Cond(j).Caption
                    .ROW = SpreadHeader + 1
                    .Text = chk_Cond(j).Caption
                    
                    .Col = .MaxCols - cur_row
                    .ROW = SpreadHeader
                    .Text = "钢种名称"
                    .ROW = SpreadHeader + 1
                    .Text = "钢种名称"
                    
                    .Col = .MaxCols - cur_row - 1: .Col2 = .MaxCols - cur_row - 1
                    .ROW = SpreadHeader: .Row2 = SpreadHeader + 1
                    .ColMerge = MergeRestricted
                    
                    .Col = .MaxCols - cur_row: .Col2 = .MaxCols - cur_row
                    .ROW = SpreadHeader: .Row2 = SpreadHeader + 1
                    .ColMerge = MergeRestricted
               Else
                    .MaxCols = .MaxCols + 1
                    .InsertCols .MaxCols - cur_row, 1
                    .Col = .MaxCols - cur_row
                    .ROW = SpreadHeader
                    .Text = chk_Cond(j).Caption
                    .ROW = SpreadHeader + 1
                    .Text = chk_Cond(j).Caption
                    
                    .Col = .MaxCols - cur_row: .Col2 = .MaxCols - cur_row
                    .ROW = SpreadHeader: .Row2 = SpreadHeader + 1
                    .ColMerge = MergeRestricted
               End If
            End If
        Next j
        
        .Refresh
        
    End With
    
    Exit Function

SpreadDisplay_Error:
    
    Sp_Header_Refer = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay_Error : " & Error)
    
End Function

Public Function Sp_Data_Refer(sPname) As Boolean

On Error GoTo SpreadDisplay_Error

    Dim iCol As Integer
    Dim iRow As Integer
    Dim iCnt As Integer
    Dim sTdate As String
    Dim sQuery As String
    Dim sEdate, sEdate1, sEdate2 As String
    Dim sTplt_prc As String
    Dim sTprc_line As String
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant
    Dim i As Integer, j As Integer

    Set AdoRs = New ADODB.Recordset
    
    sQuery = "{CALL CGT2001C.P_SREFER('" + ud_ch_date1.RawData + "','" + ud_ch_date2.RawData + "', '" + Trim(txt_cool(1).Text) + "','" + Trim(txt_cool(2).Text) + "', '" + Trim(txt_cool(3).Text) + "','" + txt_groupby.Text + "')}"
    
    'Ado Execute
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    With sPname

        Sp_Data_Refer = True
        .ReDraw = False
        Screen.MousePointer = vbHourglass
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            Sp_Data_Refer = False
            .ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 2) + 1 <> 0 Then
            .MaxRows = UBound(ArrayRecords, 2) + 1
            For iCnt = 0 To UBound(ArrayRecords, 2)
                .ROW = iCnt + 1
                For iCol = 1 To .MaxCols
                    .Col = iCol
                    .Text = IIf(IsNull(Trim(ArrayRecords(iCol - 1, iCnt))), "", Trim(ArrayRecords(iCol - 1, iCnt)))
                    .TypeHAlign = TypeHAlignRight
                    .TypeVAlign = TypeVAlignCenter
                    .Lock = True
                Next iCol
            Next iCnt
        End If
        
        For i = 1 To 6
            .Col = i
            .ROW = SpreadHeader
            If .Text = chk_Cond(1).Caption Then
               For j = 1 To .MaxRows
                    .ROW = j
                    .TypeHAlign = TypeHAlignCenter
               Next j
               .Col = i + 1
               For j = 1 To .MaxRows
                    .ROW = j
                    .TypeHAlign = TypeHAlignLeft
               Next j
            End If
            
            If .Text = chk_Cond(5).Caption Then
               For j = 1 To .MaxRows
                    .ROW = j
                    .TypeHAlign = TypeHAlignLeft
               Next j
            End If
        Next i
        
        .ReDraw = True
        Screen.MousePointer = vbDefault
        
    End With
    
    MDIMain.StatusBar1.Panels(1) = "提示信息: 数据查询完成"
    Exit Function

SpreadDisplay_Error:
    
    Set AdoRs = Nothing
    Sp_Data_Refer = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay_Error : " & Error)
    
End Function
