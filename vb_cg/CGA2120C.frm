VERSION 5.00
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form CGA2120C 
   Caption         =   "板坯产出实绩查询界面_CGA2120C"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_SlabNo 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   1200
      MaxLength       =   10
      TabIndex        =   0
      Tag             =   "EMP_ID"
      Top             =   180
      Width           =   1395
   End
   Begin VB.ComboBox cbo_Shift 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "CGA2120C.frx":0000
      Left            =   12330
      List            =   "CGA2120C.frx":000D
      TabIndex        =   3
      Top             =   150
      Width           =   810
   End
   Begin VB.ComboBox cbo_Group 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "CGA2120C.frx":001A
      Left            =   14010
      List            =   "CGA2120C.frx":002A
      TabIndex        =   4
      Top             =   150
      Width           =   780
   End
   Begin InDate.UDate txt_Occr_Time 
      Height          =   315
      Left            =   3810
      TabIndex        =   1
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
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   13200
      Top             =   150
      Width           =   750
      _ExtentX        =   1323
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
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   11550
      Top             =   150
      Width           =   750
      _ExtentX        =   1323
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
   Begin FPSpread.vaSpread ss1 
      Height          =   8115
      Left            =   90
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   630
      Width           =   14925
      _Version        =   393216
      _ExtentX        =   26326
      _ExtentY        =   14314
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
      MaxCols         =   28
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "CGA2120C.frx":003A
   End
   Begin InDate.ULabel 板坯号 
      Height          =   315
      Left            =   90
      Top             =   180
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   556
      Caption         =   "板坯号"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   2640
      Top             =   180
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   556
      Caption         =   "生产时间"
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
   Begin InDate.UDate txt_Occr_Time_To 
      Height          =   315
      Left            =   5520
      TabIndex        =   2
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
   Begin InDate.UDate txt_mill_Time 
      Height          =   315
      Left            =   8220
      TabIndex        =   7
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
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   7050
      Top             =   150
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   556
      Caption         =   "轧制时间"
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
   Begin InDate.UDate txt_mill_to_Time 
      Height          =   315
      Left            =   9930
      TabIndex        =   8
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
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "~~"
      Height          =   120
      Left            =   9690
      TabIndex        =   9
      Top             =   270
      Width           =   135
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "~~"
      Height          =   120
      Left            =   5280
      TabIndex        =   6
      Top             =   300
      Width           =   135
   End
End
Attribute VB_Name = "CGA2120C"
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
'-- Program Name      板坯产出实绩查询
'-- Program ID
'-- Document No       Q-00-0010(Specification)
'-- Designer          HANCHAO
'-- Coder             HANCHAO
'-- Date              2013.4.30
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

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
          Call Gp_Ms_Collection(txt_SlabNo, "p", " ", " ", " ", "", " ", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_Occr_Time, "p", " ", " ", " ", "", " ", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_Occr_Time_To, "p", " ", " ", " ", "", " ", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(CBO_SHIFT, "p", " ", " ", " ", "", " ", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(CBO_GROUP, "p", " ", " ", " ", "", " ", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_mill_Time, "p", " ", " ", " ", "", " ", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_mill_to_Time, "p", " ", " ", " ", "", " ", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               
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
    Call Gp_Sp_Collection(ss1, 1, "", "", "", "", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, "", "", "", "", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, "", " ", " ", "", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, "", " ", " ", "", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, "", " ", " ", "", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, "", " ", " ", "", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, "", " ", " ", " ", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, "", " ", " ", " ", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, "", " ", " ", "", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 10, "", " ", " ", "", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 11, "", "", "", "", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 12, "", "", "", "", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 13, "", " ", " ", "", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 14, "", " ", " ", "", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 15, "", " ", " ", "", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 16, "", " ", " ", "", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 17, "", " ", " ", " ", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 18, "", " ", " ", " ", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 19, "", " ", " ", " ", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 20, "", " ", " ", " ", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 21, "", " ", " ", " ", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 22, "", " ", " ", "", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 23, "", " ", " ", "", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 24, "", "", "", "", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 25, "", "", "", "", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 26, "", "", "", "", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 27, "", "", "", "", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 28, "", "", "", "", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
'    sc1.Add Item:="CGG9999C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="CGA2120C.P_SREFER", Key:="P-R"
'    sc1.Add Item:="CGG9999C.P_ONEROW", Key:="P-O"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
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
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "Z-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault
    
    txt_Occr_Time.RawData = Format(Now, "YYYYMMDD")
    txt_Occr_Time_To.RawData = Format(Now, "YYYYMMDD")

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "Z-System.INI", Me.Name)
    
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

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        rControl(1).SetFocus
    End If

End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err
    
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Exit Sub
    End If
    
    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Pro()

    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    
End Sub

Public Sub Form_Ins()
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))

End Sub

Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    
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

Public Sub Spread_Del()
    
    Call Gp_Sp_Del(Proc_Sc("SC"))

End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)
    
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, ROW)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal ROW As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
    End If
    
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub

    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
     Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
   End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

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
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub


