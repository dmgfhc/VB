VERSION 5.00
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form CGD2062C 
   Caption         =   "探伤日报表查询_CGD2062C"
   ClientHeight    =   10680
   ClientLeft      =   15
   ClientTop       =   1740
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10680
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin FPSpread.vaSpread ss1 
      Height          =   8505
      Left            =   105
      TabIndex        =   0
      Top             =   600
      Width           =   15075
      _Version        =   393216
      _ExtentX        =   26591
      _ExtentY        =   15002
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   16
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "CGD2062C.frx":0000
   End
   Begin Threed.SSFrame Single 
      Height          =   690
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   1217
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
      Begin VB.TextBox TXT_OVER_FL 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   13800
         MaxLength       =   2
         TabIndex        =   8
         Tag             =   "标准代码"
         Top             =   120
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.TextBox TXT_CO_CD 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   14400
         MaxLength       =   2
         TabIndex        =   5
         Tag             =   "标准代码"
         Top             =   120
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.CheckBox chk_Cond_W 
         BackColor       =   &H00E0E0E0&
         Caption         =   "协力"
         Height          =   255
         Left            =   12720
         TabIndex        =   4
         Tag             =   "W"
         Top             =   120
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.CheckBox chk_Cond_B 
         BackColor       =   &H00E0E0E0&
         Caption         =   "板卷"
         Height          =   255
         Left            =   12720
         TabIndex        =   3
         Tag             =   "B"
         Top             =   360
         Visible         =   0   'False
         Width           =   990
      End
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Left            =   555
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         Caption         =   "探伤日期"
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
      Begin InDate.UDate SDT_PROD_DATE_FROM 
         Height          =   315
         Left            =   1860
         TabIndex        =   7
         Tag             =   "起始日期"
         Top             =   180
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
      Begin InDate.UDate SDT_PROD_DATE_TO 
         Height          =   315
         Left            =   3600
         TabIndex        =   6
         Tag             =   "起始日期"
         Top             =   180
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
      Begin Threed.SSOption OPT_OVER 
         Height          =   330
         Left            =   6570
         TabIndex        =   9
         Top             =   210
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
         Caption         =   "异常坯"
      End
      Begin Threed.SSOption OPT_ALL 
         Height          =   330
         Left            =   5520
         TabIndex        =   10
         Top             =   210
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
         Caption         =   "全部"
      End
      Begin Threed.SSOption OPT_NOT_OVER 
         Height          =   330
         Left            =   7680
         TabIndex        =   11
         Top             =   210
         Width           =   1185
         _ExtentX        =   2090
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
         Caption         =   "非异常坯"
      End
      Begin Threed.SSOption OPT_HEAD 
         Height          =   330
         Left            =   9000
         TabIndex        =   12
         Top             =   210
         Width           =   825
         _ExtentX        =   1455
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
         Caption         =   "头坯"
      End
      Begin Threed.SSOption OPT_TAIL 
         Height          =   330
         Left            =   10080
         TabIndex        =   13
         Top             =   210
         Width           =   1185
         _ExtentX        =   2090
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
         Caption         =   "尾坯"
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "~"
         Height          =   180
         Left            =   3390
         TabIndex        =   2
         Top             =   300
         Width           =   120
      End
   End
End
Attribute VB_Name = "CGD2062C"
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
'-- Program Name      探伤日报表查询界面
'-- Program ID        CGD2062C
'-- Document No       Q-00-0010(Specification)
'-- Designer          KIM.S.H
'-- Coder             KIM.S.H
'-- Date              2008.09.22
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

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Refer"

     'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
   Call Gp_Ms_Collection(SDT_PROD_DATE_FROM, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(SDT_PROD_DATE_TO, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(TXT_CO_CD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_OVER_FL, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"

     Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    'Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
   
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="CGD2062C.P_SREFER", Key:="P-R"
    sc1.Add Item:=pColumn, Key:="pColumn"
    sc1.Add Item:=nColumn, Key:="nColumn"
    sc1.Add Item:=aColumn, Key:="aColumn"
    sc1.Add Item:=mColumn, Key:="mColumn"
    sc1.Add Item:=iColumn, Key:="iColumn"
    sc1.Add Item:=lColumn, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"

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

    Call Gf_Sp_Cls(Proc_Sc("Sc"))

    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "G-System.INI", Me.Name)
    
    SDT_PROD_DATE_FROM.RawData = Gf_DTSet(M_CN1, "D")
    SDT_PROD_DATE_TO.RawData = Gf_DTSet(M_CN1, "D")
    
    opt_all.Value = True
    
    If opt_all.Value = True Then
        TXT_OVER_FL.Text = "1"
    End If
    
    If OPT_OVER.Value Then
        TXT_OVER_FL.Text = "2"
    End If
    
    If OPT_NOT_OVER.Value Then
      TXT_OVER_FL.Text = "3"
    End If
    
    If OPT_HEAD.Value Then
      TXT_OVER_FL.Text = "4"
    End If
    
    If OPT_TAIL.Value Then
      TXT_OVER_FL.Text = "5"
    End If
        
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "G-System.INI", Me.Name)
    
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

Public Sub Form_Exit()

    Unload Me

End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(sc1) Then
       Call Gp_Ms_Cls(Mc1("rControl"))
       Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
       Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    End If

End Sub

Public Sub Master_Cpy()

    Call Gf_Ms_Copy(Mc1)

End Sub

Public Sub Master_Pst()

     If Gf_Ms_Paste(M_CN1, Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
     End If

End Sub

Public Sub Form_Ref()
    
    If Not Gp_DateCheck(SDT_PROD_DATE_FROM.Text, "S") Or Not Gp_DateCheck(SDT_PROD_DATE_TO.Text, "S") Then
       Call Gp_MsgBoxDisplay("请正确输入时间..")
       Exit Sub
    End If
                        
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        Call Data_Sum_Edit
        ss1.OperationMode = OperationModeNormal
'        Call Zero_Cls
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
               
End Sub

Public Sub Zero_Cls()
    Dim iRow  As Long
    Dim iCol  As Long
    
    For iRow = 1 To ss1.MaxRows
        ss1.ROW = iRow
        For iCol = 3 To ss1.MaxCols
            ss1.Col = iCol
            If Val(ss1.Text & "") = 0 Then
                ss1.Text = ""
            End If
        Next iCol
    Next iRow
End Sub

Public Sub Form_Pro()

     If Gf_Mc_Authority(sAuthority, Mc1) Then
       ' txt_ins_emp.Text = sUserID
       If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
    End If

End Sub

Public Sub Form_Del()

    If Not Gf_Ms_Del(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

End Sub

Public Sub Form_Exc()
    
    Call ExcelPrn
End Sub

Private Sub Data_Sum_Edit()
    Dim cSum()      As Double
    Dim cSumTotal() As Double
    Dim sSpecTemp   As String
    Dim dThkTemp    As Double
    Dim sSpec       As String
    Dim dThk        As Double
    Dim iIdr        As Integer
    Dim iIdc        As Integer
    Dim iRow        As Integer
    
    ReDim cSum(3 To 14)
    ReDim cSumTotal(3 To 14)
    
    iRow = 0
    With ss1
        For iIdr = 1 To .MaxRows
            iRow = iRow + 1
            
            .ROW = iRow
            .Col = 1
            sSpec = .Text
            ss1.Col = 2
            dThk = Val(.Text & "")
            
            If sSpecTemp <> sSpec And iRow <> 1 Then
                .MaxRows = .MaxRows + 1
                .InsertRows iRow, 1
                .ROW = iRow
                .Col = 1:    .Text = sSpecTemp
                .Col = 2:    .Text = "小计"
                
                For iIdc = 3 To 14
                    .Col = iIdc
                    .Text = cSum(iIdc)
                    cSum(iIdc) = 0
                Next iIdc
                
                iIdr = iIdr - 1
            Else
                For iIdc = 3 To 14
                    .ROW = iRow
                    .Col = iIdc
                    cSum(iIdc) = cSum(iIdc) + Val(.Text & "")
                    cSumTotal(iIdc) = cSumTotal(iIdc) + Val(.Text & "")
                Next iIdc
            End If
                
            sSpecTemp = sSpec
            dThkTemp = dThk
        Next iIdr
        
        .MaxRows = .MaxRows + 2
        For iIdc = 3 To 14
            .ROW = .MaxRows - 1
            .Col = 1:    .Text = sSpecTemp
            .Col = 2:    .Text = "小计"
            .Col = iIdc: .Text = cSum(iIdc)
            
            .ROW = .MaxRows
            .Col = 1:    .Text = "合计"
            .Col = iIdc: .Text = cSumTotal(iIdc)
        Next iIdc
        
        ReDim cSum(1 To 6)
        
        For iIdr = 1 To .MaxRows
            .ROW = iIdr
            .Col = 3:    cSum(1) = Val(.Text & "")
            .Col = 4:    cSum(2) = Val(.Text & "")
            .Col = 7:    cSum(3) = Val(.Text & "")
            .Col = 8:    cSum(4) = Val(.Text & "")
            .Col = 11:   cSum(5) = Val(.Text & "")
            .Col = 12:   cSum(6) = Val(.Text & "")
            If cSum(5) > 0 And cSum(1) > 0 Then .Col = 15: .Text = cSum(1) / cSum(5) * 100
            If cSum(6) > 0 And cSum(2) > 0 Then .Col = 16: .Text = cSum(2) / cSum(6) * 100
            'If cSum(6) > 0 Then .Col = 17:   .Text = cSum(3) / cSum(6) * 100
        Next iIdr
        
    End With

End Sub

Private Sub ExcelPrn()
    Dim i               As Integer
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

    xlApp.Workbooks.Open (App.Path & "\CGD2062C.xls")

    Set xlSheet = xlApp.Worksheets("Sheet1")
    xlApp.Sheets("Sheet1").Select
    
    For i = 2 To ss1.MaxRows
          xlApp.Rows("4:4").Select
          xlApp.Selection.Copy
          xlApp.Selection.Insert Shift:=1
    Next i
            
    sDateFr = SDT_PROD_DATE_FROM.Text
    sDateTo = SDT_PROD_DATE_TO.Text

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
Private Sub chk_Cond_B_Click()

    If chk_Cond_B Then
        TXT_CO_CD.Text = chk_Cond_B.Tag
        chk_Cond_W = False
    End If
    
    If chk_Cond_B = False And chk_Cond_W = False Then
        TXT_CO_CD.Text = ""
    End If
    
End Sub

Private Sub chk_Cond_W_Click()

    If chk_Cond_W Then
        TXT_CO_CD.Text = chk_Cond_W.Tag
        chk_Cond_B = False
    End If
    
    If chk_Cond_B = False And chk_Cond_W = False Then
        TXT_CO_CD.Text = ""
    End If
    
End Sub

Private Sub opt_all_Click(Value As Integer)
    If opt_all.Value = True Then
        opt_all.ForeColor = &HFF&
        OPT_OVER.ForeColor = &H808080
        OPT_NOT_OVER.ForeColor = &H808080
        OPT_HEAD.ForeColor = &H808080
        OPT_TAIL.ForeColor = &H808080
        TXT_OVER_FL.Text = "1"
    Else
        opt_all.ForeColor = &H808080
    End If
End Sub



Private Sub OPT_HEAD_Click(Value As Integer)
If OPT_HEAD.Value = True Then
        OPT_HEAD.ForeColor = &HFF&
        OPT_TAIL.ForeColor = &H808080
        OPT_NOT_OVER.ForeColor = &H808080
        OPT_OVER.ForeColor = &H808080
        opt_all.ForeColor = &H808080
        TXT_OVER_FL.Text = "4"
    Else
        OPT_HEAD.ForeColor = &H808080
    End If
End Sub

Private Sub OPT_NOT_OVER_Click(Value As Integer)
    If OPT_NOT_OVER.Value = True Then
        OPT_NOT_OVER.ForeColor = &HFF&
        OPT_OVER.ForeColor = &H808080
        opt_all.ForeColor = &H808080
        OPT_HEAD.ForeColor = &H808080
        OPT_TAIL.ForeColor = &H808080
        TXT_OVER_FL.Text = "3"
    Else
        OPT_NOT_OVER.ForeColor = &H808080
    End If
End Sub

Private Sub OPT_OVER_Click(Value As Integer)
     If OPT_OVER.Value = True Then
        OPT_OVER.ForeColor = &HFF&
        opt_all.ForeColor = &H808080
        OPT_HEAD.ForeColor = &H808080
        OPT_TAIL.ForeColor = &H808080
        OPT_NOT_OVER.ForeColor = &H808080
        TXT_OVER_FL.Text = "2"
    Else
        OPT_OVER.ForeColor = &H808080
    End If
End Sub

Private Sub OPT_TAIL_Click(Value As Integer)
    If OPT_TAIL.Value = True Then
        OPT_HEAD.ForeColor = &H808080
        OPT_TAIL.ForeColor = &HFF&
        OPT_NOT_OVER.ForeColor = &H808080
        OPT_OVER.ForeColor = &H808080
        opt_all.ForeColor = &H808080
        TXT_OVER_FL.Text = "5"
    Else
        OPT_TAIL.ForeColor = &H808080
    End If

End Sub

Private Sub SDT_PROD_DATE_TO_GotFocus()
     If SDT_PROD_DATE_TO.RawData = "" Then
        SDT_PROD_DATE_TO.RawData = Gf_DTSet(M_CN1, "D")
     End If
End Sub

Private Sub SDT_PROD_DATE_FROM_GotFocus()

     If SDT_PROD_DATE_FROM.RawData = "" Then
        SDT_PROD_DATE_FROM.RawData = Gf_DTSet(M_CN1, "D")
     End If
     
     If SDT_PROD_DATE_TO.RawData = "" Then
        SDT_PROD_DATE_TO.RawData = Gf_DTSet(M_CN1, "D")
     End If
     
End Sub



