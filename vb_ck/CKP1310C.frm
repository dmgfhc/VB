VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form CKP1310C 
   Caption         =   "中板厂生产日报_CKP1310C"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter AW 
      Height          =   8670
      Left            =   105
      TabIndex        =   0
      Top             =   675
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   15293
      _Version        =   196609
      BorderStyle     =   0
      PaneTree        =   "CKP1310C.frx":0000
      Begin FPSpread.vaSpread ss3 
         Height          =   3480
         Left            =   0
         TabIndex        =   6
         Top             =   5190
         Width           =   15120
         _Version        =   393216
         _ExtentX        =   26670
         _ExtentY        =   6138
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
         MaxCols         =   15
         MaxRows         =   7
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CKP1310C.frx":0072
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   2280
         Left            =   0
         TabIndex        =   5
         Top             =   2820
         Width           =   15120
         _Version        =   393216
         _ExtentX        =   26670
         _ExtentY        =   4022
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
         MaxCols         =   8
         MaxRows         =   6
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CKP1310C.frx":1EB9
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   2730
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   15120
         _Version        =   393216
         _ExtentX        =   26670
         _ExtentY        =   4815
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
         MaxCols         =   15
         MaxRows         =   10
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CKP1310C.frx":2F58
      End
   End
   Begin Threed.SSFrame Single 
      Height          =   555
      Left            =   105
      TabIndex        =   1
      Top             =   90
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   979
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
      Begin Threed.SSCommand Cmd_Edit 
         Height          =   360
         Left            =   10335
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   90
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   635
         _Version        =   196609
         Font3D          =   1
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "更新数据"
      End
      Begin InDate.UDate txt_DATE 
         Height          =   315
         Left            =   2595
         TabIndex        =   3
         Tag             =   "起始日期"
         Top             =   120
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
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Left            =   1410
         Top             =   120
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         Caption         =   "日期"
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
   End
End
Attribute VB_Name = "CKP1310C"
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
'-- Program Name      PROD REPORT
'-- Program ID        CKP1310C
'-- Designer          GUOLI
'-- Coder             GUOLI
'-- Date              2008.08.13
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

Dim pColumn2 As New Collection      'Spread Primary Key Collection
Dim nColumn2 As New Collection      'Spread necessary Column Collection
Dim mColumn2 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn2 As New Collection      'Spread Insert Column Collection
Dim aColumn2 As New Collection      'Master -> Spread Column Collection
Dim lColumn2 As New Collection      'Spread Lock Column Collection

Dim pColumn3 As New Collection      'Spread Primary Key Collection
Dim nColumn3 As New Collection      'Spread necessary Column Collection
Dim mColumn3 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn3 As New Collection      'Spread Insert Column Collection
Dim aColumn3 As New Collection      'Master -> Spread Column Collection
Dim lColumn3 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim Sc1 As New Collection           'Spread Collection
Dim Sc2 As New Collection           'Spread Collection
Dim Sc3 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
    Dim i As Integer
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Sheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
    Call Gp_Ms_Collection(txt_DATE, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

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
    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="CKP1310C.P_SREFER1", Key:="P-R"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxRows, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc1"

    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    
    'Spread_Collection
    Sc2.Add Item:=ss2, Key:="Spread"
    Sc2.Add Item:="CKP1310C.P_SREFER2", Key:="P-R"
    Sc2.Add Item:=pColumn2, Key:="pColumn"
    Sc2.Add Item:=nColumn2, Key:="nColumn"
    Sc2.Add Item:=aColumn2, Key:="aColumn"
    Sc2.Add Item:=mColumn2, Key:="mColumn"
    Sc2.Add Item:=iColumn2, Key:="iColumn"
    Sc2.Add Item:=lColumn2, Key:="lColumn"
    Sc2.Add Item:=1, Key:="First"
    Sc2.Add Item:=ss2.MaxRows, Key:="Last"
    Proc_Sc.Add Item:=Sc2, Key:="Sc2"
    
        'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss3, 1, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 2, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 3, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 4, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 5, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 6, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 7, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 8, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 9, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 10, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 11, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 12, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 13, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 14, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 15, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    
    'Spread_Collection
    Sc3.Add Item:=ss3, Key:="Spread"
    Sc3.Add Item:="CKP1310C.P_SREFER3", Key:="P-R"
    Sc3.Add Item:=pColumn3, Key:="pColumn"
    Sc3.Add Item:=nColumn3, Key:="nColumn"
    Sc3.Add Item:=aColumn3, Key:="aColumn"
    Sc3.Add Item:=mColumn3, Key:="mColumn"
    Sc3.Add Item:=iColumn3, Key:="iColumn"
    Sc3.Add Item:=lColumn3, Key:="lColumn"
    Sc3.Add Item:=1, Key:="First"
    Sc3.Add Item:=ss1.MaxRows, Key:="Last"
    Proc_Sc.Add Item:=Sc3, Key:="Sc3"

'    ss3.Col = 2: ss3.Col2 = 2
'    ss3.ROW = 49: ss3.Row2 = 49
'
'    ss3.Lock = False
'    ss3.BlockMode = False
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)

End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    Call Form_Define
        
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Sp_Setting(Sc1.Item("Spread"))
    Call Sp_Setting(Sc2.Item("Spread"))
    Call Sp_Setting(Sc3.Item("Spread"))
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gp_Sp_ColGet(Proc_Sc("Sc1")("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc2")("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc3")("Spread"), "K-System.INI", Me.Name)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
       Cmd_Edit.Enabled = True
    End If

    txt_DATE.RawData = Format(Date - 1, "yyyymmdd")
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("Sc1")("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc2")("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc3")("Spread"), "K-System.INI", Me.Name)
    
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
   
    Set iColumn3 = Nothing
    Set pColumn3 = Nothing
    Set lColumn3 = Nothing
    Set nColumn3 = Nothing
    Set mColumn3 = Nothing
    Set aColumn3 = Nothing
   
    Set Mc1 = Nothing
    Set Sc1 = Nothing
    Set Sc2 = Nothing
    Set Sc3 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
End Sub

Public Sub Form_Cls()
    Dim iRow  As Long
    Dim iCol  As Long
    
    Call Form_SP_Cls
    
'    Call Gf_Sp_Cls(Proc_Sc("Sc3"))
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_ControlLock(Mc1("lControl"), False)

    txt_DATE.RawData = Format(Date - 1, "yyyymmdd")
End Sub

Public Sub Form_SP_Cls()
    Dim iRow  As Long
    Dim iCol  As Long
    
    For iRow = 1 To ss1.MaxRows
        ss1.ROW = iRow
        For iCol = 1 To ss1.MaxCols
           ss1.Col = iCol
           ss1.Text = ""
        Next iCol
    Next iRow
    
    For iRow = 1 To ss2.MaxRows
        ss2.ROW = iRow
        For iCol = 1 To ss2.MaxCols
           ss2.Col = iCol
           ss2.Text = ""
        Next iCol
    Next iRow
    
    For iRow = 1 To ss3.MaxRows
         ss3.ROW = iRow
         For iCol = 3 To ss3.MaxCols
             ss3.Col = iCol
             If ss3.CellType = SS_CELL_TYPE_NUMBER Then
                ss3.Text = ""
             End If
         Next iCol
    Next iRow
End Sub

Public Sub Form_Ref()
    
    If Trim(txt_DATE.Text) = "" Then
        Call Gp_MsgBoxDisplay(txt_DATE.Tag + "必须输入")
        Exit Sub
    End If
    
    Call Form_SP_Cls
    Screen.MousePointer = vbHourglass
    
    If Sp_Display(M_CN1, Proc_Sc("Sc1")("Spread"), Gf_Ms_MakeQuery(Proc_Sc("Sc1").Item("P-R"), "R", Mc1("pControl"))) Then
       Call Sp_Display2(M_CN1, Proc_Sc("Sc2")("Spread"), Gf_Ms_MakeQuery(Proc_Sc("Sc2").Item("P-R"), "R", Mc1("pControl")))
       Call Sp_Display3(M_CN1, Proc_Sc("Sc3")("Spread"), Gf_Ms_MakeQuery(Proc_Sc("Sc3").Item("P-R"), "R", Mc1("pControl")))
       Call SearchCommentsData
    End If
    
    ss1.ReDraw = True
    ss2.ReDraw = True
    ss3.ReDraw = True
    
    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    
    Screen.MousePointer = vbDefault
    
End Sub

Public Sub Form_Exc()

'    Call Gp_Sp_Excel(Me, Proc_Sc("Sc1")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    Call ExcelPrn
    
End Sub

Public Sub Form_Pro()
    Dim sQuery      As String
    Dim sComments   As String
    Dim sDate       As String
    
    On Error GoTo UPDATE_ERROR

    Screen.MousePointer = vbHourglass
    
    M_CN1.BeginTrans
 
    ss3.ROW = 49
    ss3.Col = 2
    sComments = Trim(ss3.Text)
    
    sDate = Format(txt_DATE.Text, "YYYYMMDD")
    
    sQuery = ""
    sQuery = "         UPDATE  gp_zbrpt_mon                                 " & vbCrLf
    sQuery = sQuery & "   SET  COMMENT1         = '" & sComments & "'       " & vbCrLf
    sQuery = sQuery & " WHERE  PLT              = 'C3'                      " & vbCrLf
    sQuery = sQuery & "   AND  PROD_DATE        = '" & sDate & "'           " & vbCrLf

    M_CN1.Execute sQuery
        
    M_CN1.CommitTrans

    Screen.MousePointer = vbDefault
    Exit Sub

UPDATE_ERROR:

    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay(Err.Description & sQuery)
    
    M_CN1.RollbackTrans
    
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

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)
    
    'Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Sub SearchCommentsData()

    Dim AdoRs As New ADODB.Recordset
    Dim sql               As String
    Dim sDate             As String
    Dim i, j              As Integer
    
    sDate = Format(txt_DATE.Text, "YYYYMMDD")
    
    sql = "      SELECT  YEAR_PLAN_WGT   ,   YEAR_FIN_WGT    , YEAR_PROD_PROG   ,  " & vbCrLf
    sql = sql + "        YEAR_PROG       ,   YEAR_AVE_NEED   , YEAR_LEFT_DAY    ,  " & vbCrLf
    sql = sql + "        NULL            ,                                         " & vbCrLf
    sql = sql + "        MONTH_PLAN_WGT  ,   MONTH_FIN_WGT   , MONTH_PROD_PROG  ,  " & vbCrLf
    sql = sql + "        MONTH_PROG      ,   MONTH_AVE_NEED  , MONTH_LEFT_DAY   ,  " & vbCrLf
    sql = sql + "        NULL                                                      " & vbCrLf
    sql = sql & "  FROM  gp_zb_mon                                              " & vbCrLf
    sql = sql & " WHERE  PLT                     = 'C3'                            " & vbCrLf
    sql = sql & "   AND  PROD_DATE               = '" & sDate & "'                 " & vbCrLf
    
    AdoRs.Open sql, M_CN1, adOpenForwardOnly, adLockReadOnly
    If Not AdoRs.EOF Then
       With ss3
            For j = 6 To 7
                .ROW = j
                For i = 2 To 14 Step 2
                    .Col = i
                    If j = 6 Then
                        If Not (VarType(AdoRs.Fields(i / 2 - 1)) = vbNull Or AdoRs.Fields(i / 2 - 1).Value = 0) Then
                          .Text = Val(AdoRs.Fields(i / 2 - 1))
                        End If
                    Else
                        If Not (VarType(AdoRs.Fields(i / 2 + 6)) = vbNull Or AdoRs.Fields(i / 2 + 6).Value = 0) Then
                          .Text = Val(AdoRs.Fields(i / 2 + 6))
                        End If
                    End If
                Next
            Next
        End With
    End If
    
    AdoRs.Close
    
End Sub
Private Sub Cmd_Edit_Click()
    'On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim strRet_Result_ErrMsg As String
    Dim sQuery As String
          
    If Trim(txt_DATE.Text) = "" Then
        Call Gp_MsgBoxDisplay(txt_DATE.Tag + "必须输入")
        Exit Sub
    End If

    Dim adoCmd As ADODB.Command
    
     Screen.MousePointer = vbHourglass
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    sQuery = "{call CKP1310P ('" + Trim(Format(txt_DATE.Text, "YYYYMMDD")) + "',?)}"

    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
            
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    
    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    If adoCmd("arg_e_msg") <> "" Then
        strRet_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & strRet_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
    Else
        Set adoCmd = Nothing
        Screen.MousePointer = vbDefault
        
        Call Gp_MsgBoxDisplay("更新成功..!!", "I")
        Call Form_Ref
        Exit Sub
    End If
Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("更新失败！！")

End Sub


Private Sub ExcelPrn()
    Dim i               As Integer
    Dim xlApp           As Object
    Dim xlSheet         As Object
    Dim sDate           As String
    
    If ss1.MaxRows < 1 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
     
    On Error Resume Next
    
    Set xlApp = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then
        Set xlApp = CreateObject("Excel.Application")
    End If
    
    Err.Clear

    xlApp.Workbooks.Open (App.Path & "\CKP1310C.xls")
    
    Set xlSheet = xlApp.Worksheets("Sheet1")
    xlApp.Sheets("Sheet1").Select
    
    sDate = Format(txt_DATE.Text, "YYYYMMDD")
    xlApp.Range("A2").Value = "报表日期：" + Left(sDate, 4) + "年" + Mid(sDate, 5, 2) + "月" + Mid(sDate, 7, 2) + "日"
    xlApp.Range("B35").Value = "制表日期：" + Format(Now, "YYYY-MM-DD HH:MM:SS")
    xlApp.Range("K35").Value = "制表人：" + sUserID

    Clipboard.Clear
    ss1.SetSelection 1, 1, ss1.MaxCols, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("C5").Select
    xlApp.ActiveSheet.Paste

    Clipboard.Clear
    ss2.SetSelection 1, 1, ss2.MaxCols, ss2.MaxRows
    ss2.ClipboardCopy
    xlApp.Range("C18").Select
    xlApp.ActiveSheet.Paste

    Clipboard.Clear
    ss3.SetSelection 1, 1, 14, 5
    ss3.ClipboardCopy
    xlApp.Range("B27").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss3.SetSelection 2, 6, 6, 7
    ss3.ClipboardCopy
    xlApp.Range("C33").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss3.SetSelection 8, 6, 14, 7
    ss3.ClipboardCopy
    xlApp.Range("J33").Select
    xlApp.ActiveSheet.Paste
    
    
    ss1.ClearSelection
    ss2.ClearSelection
    ss3.ClearSelection
       
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

Public Sub Sp_Setting(ByVal sPname As Variant, Optional MsgChk As Boolean = True)
    With sPname
    
        .RowHeight(-1) = 12.54
        
        If .ColHeaderRows > 1 Then
            .RowHeight(SpreadHeader + (.ColHeaderRows - 2)) = 12
            .RowHeight(SpreadHeader + (.ColHeaderRows - 1)) = 12
        Else
            .RowHeight(0) = 24
        End If
        
        .ColWidth(0) = 6
        
        .BackColorStyle = BackColorStyleUnderGrid
        
        .GrayAreaBackColor = &HE0E0E0
        .GridColor = &H808040
        
        .ShadowColor = &HE1E4CD
        .ShadowDark = &H808040
        .SelBackColor = &HCEECFF     ''&HE3F4FF      ''&HFFFF80     '&H808040
     
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
        
        If MsgChk Then
            .LockBackColor = RGB(255, 255, 255)
        End If

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
                Select Case Trim(ArrayRecords(0, iRowCount))
                    Case "A0"
                        .ROW = 1
                        ss3.ROW = 1
                    Case "A1"
                        .ROW = 2
                    Case "B0"
                        ss3.ROW = 2
                        .ROW = 3
                    Case "B1"
                        .ROW = 4
                    Case "C0"
                        ss3.ROW = 3
                        .ROW = 5
                    Case "C1"
                        .ROW = 6
                    Case "D0"
                        ss3.ROW = 4
                        .ROW = 7
                    Case "D1"
                        .ROW = 8
                    Case "T0"
                        ss3.ROW = 5
                        .ROW = 9
                    Case "T1"
                        .ROW = 10
                End Select
            
'            .ROW = iRowCount + 1

                For iColcount = 1 To .MaxCols
    
                    .Col = iColcount
    
                    If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Or ArrayRecords(iColcount, iRowCount) = 0 Then
                        .Text = ""
                    Else
                        .Text = Trim(ArrayRecords(iColcount, iRowCount))
                    End If

                Next iColcount

            Next iRowCount
            
'            For iRowCount = 0 To ss3.MaxRows - 1
'            With ss3
'                    Select Case Trim(ArrayRecords(0, iRowCount))
'                            Case "A0"
'                                .ROW = 1
'                                ss3.ROW = 1
'                            Case "A1"
'                                .ROW = 2
'                            Case "B0"
'                                ss3.ROW = 2
'                                .ROW = 3
'                            Case "B1"
'                                .ROW = 4
'                            Case "C0"
'                                ss3.ROW = 3
'                                .ROW = 5
'                            Case "C1"
'                                .ROW = 6
'                            Case "D0"
'                                ss3.ROW = 4
'                                .ROW = 7
'                            Case "D1"
'                                .ROW = 8
'                            Case "T0"
'                                ss3.ROW = 5
'                                .ROW = 9
'                            Case "T1"
'                                .ROW = 10
'                    End Select
'            End With
        End If

        .ReDraw = True
        Screen.MousePointer = vbDefault

    End With

End Function

Public Function Sp_Display2(Conn As ADODB.Connection, sPname As Variant, sQuery As String) As Boolean

    On Error Resume Next

    Dim iCount          As Integer
    Dim iRowCount       As Long
    Dim iColcount       As Long
    Dim AdoRs           As ADODB.Recordset
    Dim ArrayRecords    As Variant

    Sp_Display2 = True

    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then Sp_Display2 = False: Exit Function
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
            Sp_Display2 = False
            Call Gp_MsgBoxDisplay("无相关记录", "I")
            Call Form_Cls
            Screen.MousePointer = vbDefault
            Exit Function

        End If
        
        ArrayRecords = AdoRs.GetRows
        
        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 2) <> 0 Then

            For iRowCount = 0 To UBound(ArrayRecords, 2)

                For iColcount = 1 To 24
                    If iColcount >= 1 And iColcount < 9 Then
                       .Col = iColcount
                       .ROW = iRowCount + 1
                    ElseIf iColcount >= 9 And iColcount < 17 Then
                       .Col = iColcount - 8
                       .ROW = iRowCount + 3
                    ElseIf iColcount >= 17 And iColcount < 25 Then
                       .Col = iColcount - 16
                       .ROW = iRowCount + 5
                    End If
    
                    If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Or ArrayRecords(iColcount, iRowCount) = 0 Then
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

Public Function Sp_Display3(Conn As ADODB.Connection, sPname As Variant, sQuery As String) As Boolean

    On Error Resume Next

    Dim iCount          As Integer
    Dim iRowCount       As Long
    Dim iColcount       As Long
    Dim AdoRs           As ADODB.Recordset
    Dim ArrayRecords    As Variant

    Sp_Display3 = True

    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then Sp_Display3 = False: Exit Function
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
            Sp_Display3 = False
            Call Gp_MsgBoxDisplay("无相关记录", "I")
            Call Form_Cls
            Screen.MousePointer = vbDefault
            Exit Function

        End If
        
        ArrayRecords = AdoRs.GetRows
        
        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 2) <> 0 Then
        

            For iRowCount = 0 To UBound(ArrayRecords, 2)
                
                For iColcount = 1 To 7
                    If Mid(ArrayRecords(0, iRowCount), 2, 1) = "0" Then
                      .Col = 2 * iColcount - 1
                    ElseIf Mid(ArrayRecords(0, iRowCount), 2, 1) = "1" Then
                      .Col = 2 * iColcount
                    End If
                    
                    If Mid(ArrayRecords(0, iRowCount), 1, 1) = "A" Then
                      .ROW = 1
                    ElseIf Mid(ArrayRecords(0, iRowCount), 1, 1) = "B" Then
                      .ROW = 2
                    ElseIf Mid(ArrayRecords(0, iRowCount), 1, 1) = "C" Then
                      .ROW = 3
                    ElseIf Mid(ArrayRecords(0, iRowCount), 1, 1) = "D" Then
                      .ROW = 4
                    ElseIf Mid(ArrayRecords(0, iRowCount), 1, 1) = "T" Then
                      .ROW = 5
                    End If
    
                    If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Or ArrayRecords(iColcount, iRowCount) = 0 Then
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

