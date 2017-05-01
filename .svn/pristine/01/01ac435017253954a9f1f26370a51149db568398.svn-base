VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form ACB4080C 
   Caption         =   "板坯信息查询修改_ACB4080C"
   ClientHeight    =   9225
   ClientLeft      =   525
   ClientTop       =   2055
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_len_min 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   315
      Left            =   7650
      MaxLength       =   8
      TabIndex        =   1
      Top             =   675
      Width           =   1200
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   2820
      Top             =   240
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "当前厂库"
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
   Begin FPSpread.vaSpread ss1 
      Height          =   7995
      Left            =   60
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1185
      Width           =   15225
      _Version        =   393216
      _ExtentX        =   26855
      _ExtentY        =   14102
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
      MaxCols         =   23
      MaxRows         =   1
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      SpreadDesigner  =   "ACB4080C.frx":0000
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   225
      Top             =   660
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "厚度"
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
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   6435
      Top             =   675
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "长度下限"
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
      Left            =   2820
      Top             =   690
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "宽度"
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
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   6435
      Top             =   240
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "标准号"
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
   Begin InDate.ULabel ULabel7 
      Height          =   315
      Left            =   8880
      Top             =   675
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "长度上限"
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
   Begin VB.Frame Frame1 
      Height          =   1065
      Left            =   120
      TabIndex        =   3
      Top             =   60
      Width           =   15105
      Begin VB.TextBox txt_stdspec_chg_ref 
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
         Left            =   7530
         MaxLength       =   18
         TabIndex        =   9
         Tag             =   "标准号"
         Top             =   180
         Width           =   2925
      End
      Begin VB.TextBox txt_slab_no 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   0
         Top             =   195
         Width           =   1200
      End
      Begin VB.TextBox txt_len_max 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   9975
         MaxLength       =   8
         TabIndex        =   8
         Top             =   615
         Width           =   1200
      End
      Begin VB.TextBox txt_wid 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   3915
         MaxLength       =   8
         TabIndex        =   7
         Top             =   630
         Width           =   1200
      End
      Begin VB.TextBox txt_thk 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   6
         Top             =   600
         Width           =   1200
      End
      Begin VB.TextBox txt_cur_inv 
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
         Left            =   4290
         TabIndex        =   5
         Top             =   180
         Width           =   1845
      End
      Begin VB.TextBox txt_cur_inv_code 
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
         Left            =   3915
         MaxLength       =   2
         TabIndex        =   4
         Top             =   180
         Width           =   375
      End
      Begin InDate.ULabel ULabel4 
         Height          =   315
         Left            =   105
         Top             =   195
         Width           =   1200
         _ExtentX        =   2117
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
   End
End
Attribute VB_Name = "ACB4080C"
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
'-- Program Name      板坯信息查询修改
'-- Program ID        ACB4080C
'-- Designer          WUTAO
'-- Coder             WUTAO
'-- Date              2006.10.26
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

'----> THIS variable USE IS ACB4080C PROGRAM <--------------
Public islab_no As Long
Public chkNo As Integer
Public sSLAB_NO As String
Public botIntLen As Long
'----------------------------------------------------------

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
            Call Gp_Ms_Collection(txt_slab_no, "p", "n", " ", "", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_cur_inv_code, "p", " ", " ", "", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_stdspec_chg_ref, "p", " ", " ", "", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_thk, "p", " ", " ", "", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_wid, "p", " ", " ", "", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_len_min, "p", " ", " ", "", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_len_max, "p", " ", " ", "", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------  BELOW EDIT ---------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    '''''ADDED BY GUOLI AT 20090108164200'''''
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    ''''''''''''''''''''''''''''''''''''''''''
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="ACB4080C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="ACB4080C.P_REFER", Key:="P-R"
    sc1.Add Item:="ACB4080C.P_ONEROW", Key:="P-O"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxRows, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").Row = 0
    sc1.Item("Spread").Text = "◎"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    Frame1.BackColor = &HE0E0E0
    Call Gp_Sp_ColHidden(ss1, 22, True)
'    cboCcm.AddItem "1"
'    cboCcm.AddItem "2"
'    cboCcm.ListIndex = 0
'
'    cboPrcLine.AddItem "1"
'    cboPrcLine.AddItem "2"
'
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
    MDIMain.MenuTool.Buttons(7).Enabled = False
    MDIMain.MenuTool.Buttons(8).Enabled = False
    MDIMain.MenuTool.Buttons(11).Enabled = False
    MDIMain.MenuTool.Buttons(12).Enabled = False

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
    
    MDIMain.MenuTool.Buttons(7).Enabled = False
    MDIMain.MenuTool.Buttons(8).Enabled = False
    MDIMain.MenuTool.Buttons(11).Enabled = False
    MDIMain.MenuTool.Buttons(12).Enabled = False

    Call Form_Activate
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)
    
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
        MDIMain.MenuTool.Buttons(7).Enabled = False
        MDIMain.MenuTool.Buttons(8).Enabled = False
        MDIMain.MenuTool.Buttons(11).Enabled = False
        MDIMain.MenuTool.Buttons(12).Enabled = False
        Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    End If
    
'    cboFromDate = ""
'    cboToDate = ""
'    cboPrcLine = ""
'    txtHeatNo = ""
'    cboShift = ""
'    cboGroup = ""
    

End Sub

Public Sub Form_Ref()
Dim i As Integer
On Error GoTo Refer_Err


    'ERROR CHECK
'    If Asc(Mid(cboFromDate, 1, 1)) < 48 Or Asc(Mid(cboFromDate, 1, 1)) > 57 Then
'       Call Gp_MsgBoxDisplay("请输入查询起始日期...!", "Q", "")
'       cboFromDate.SetFocus
'       Exit Sub
'    End If
'
'    If Asc(Mid(cboToDate, 1, 1)) < 48 Or Asc(Mid(cboToDate, 1, 1)) > 57 Then
'       Call Gp_MsgBoxDisplay("请输入查询结束日期...!", "Q", "")
'       cboToDate.SetFocus
'       Exit Sub
'    End If
       
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Nothing, Mc1("mControl")) Then
        Call Form_Activate
        ss1.SetFocus
    End If
            
    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Pro()

    If Gf_Sp_Process(M_CN1, Proc_Sc("Sc"), Mc1) Then Call Form_Activate
    
End Sub

Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1
        
End Sub

Public Sub Spread_Forzens_Setting()

    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol
    
End Sub

Public Sub Spread_Forzens_Cancel()

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

'Private Sub txt_stdspec_DblClick()
'    Call txt_stdspec_KeyUp(vbKeyF4, 0)
'End Sub

'Private Sub txt_stdspec_KeyUp(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = vbKeyF4 Then
'
'        DD.sWitch = "MS"
''        txt_stdspec_yy.Text = ""
''        DD.rControl.Add Item:=txt_stdspec_chg
''        DD.rControl.Add Item:=txt_stdspec_yy
'        DD.rControl.Add Item:=txt_stdspec_name
'        Call Gf_StdSPEC_DD(M_CN1, KeyCode)
'
'        Exit Sub
'
'    End If
'End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    'Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    If Col = 3 And Row > 0 Then
        ss1.Col = Col
        ss1.Row = Row

'        If ss1.Text = "定尺" Then
'           ss1.Text = "非定尺"
'        ElseIf ss1.Text = "非定尺" Then
'           ss1.Text = "定尺"
'        End If
        ss1.Col = 0
        ss1.Text = "Update"
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 22)
    End If
'    生产部要求将 修改定尺 非定尺 功能取消  耿学玉  2012 0509
'    If Col = 2 And Row > 0 Then
'        ss1.Col = Col
'        ss1.Row = Row
'
'        If ss1.Text = "XAC" Then
'           ss1.Text = "CAC"
'        ElseIf ss1.Text = "CAC" Then
'           ss1.Text = "XAC"
'        Else
'           ss1.Text = ss1.Text
'        End If
'        ss1.Col = 0
'        ss1.Text = "Update"
'        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 21)
'    End If
    
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 22)
    End If

End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 22)
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub

'Private Sub cmd_add_rec_Click()
'
'    Dim icount As Integer
'    Dim icount1 As Integer
'    Dim islab_no As Long
'    Dim sslab_no As String
'    Dim strTem As String
'    Dim len_z As String
'    Dim i, k, j As Integer
'    Dim A As Double
'
'    If ss1.MaxRows < 1 Then
'       Call Gp_MsgBoxDisplay("没有所要追加的纪录", "Q", "")
'       Exit Sub
'    End If
'
'    If txt_in_slab_no.Text = "" Then
'       Call Gp_MsgBoxDisplay("请输入首个外卖坯号", "Q", "")
'       Exit Sub
'    End If
'
'    For i = Len(Trim(txt_in_slab_no.Text)) To 1 Step -1
'        If Asc(Mid(txt_in_slab_no.Text, i, 1)) < 48 Or Asc(Mid(txt_in_slab_no.Text, i, 1)) > 57 Then
'          k = i
'          Exit For
'        End If
'    Next i
'
'    islab_no = CDbl(Mid(txt_in_slab_no.Text, k + 1, Len(Trim(txt_in_slab_no.Text)) - k))
'    'strTem = Mid(txt_in_slab_no.Text, 1, Len(txt_in_slab_no.Text) - k)
'   ' islab_no = txt_in_slab_no.Text
'    sslab_no = CStr(islab_no)
'    For icount = 1 To ss1.MaxRows
'        ss1.Col = 2
'        ss1.Row = icount
'        len_z = Len(txt_in_slab_no.Text) - k - Len(Trim(islab_no))
'        If len_z > 0 Then
'           For icount1 = 1 To len_z
'               sslab_no = "0" + sslab_no
'           Next icount1
'        End If
'        ss1.Text = Mid(txt_in_slab_no.Text, 1, k) + sslab_no
'        ss1.Col = 0
'        ss1.Text = "Update"
'        ss1.Col = 11
'        ss1.Text = sUserID
'        islab_no = islab_no + 1
'        sslab_no = CStr(islab_no)
'    Next icount
'
'End Sub



'Private Sub txtCurinvCd_Change()
'
'End Sub
'
'
'
'Private Sub Text1_Change()
'
'End Sub
'
'Private Sub txt_aplyspec_code_Change()
'    If Len(Trim(txt_aplyspec_code.Text)) = txt_aplyspec_code.MaxLength Then
'          txt_aplystdspec.Text = Gf_ComnNameFind(M_CN1, "G0018", txt_aplyspec_code.Text, 2)
'          Exit Sub
'    Else
'          txt_aplystdspec.Text = ""
'    End If
'End Sub
'
'Private Sub txt_aplyspec_code_KeyUp(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyF4 Then
'
'        DD.sWitch = "MS"
'        DD.sKey = "G0018"
'
'        DD.rControl.Add Item:=txt_aplyspec_code
'        DD.rControl.Add Item:=txt_aplystdspec
'
'
'        DD.nameType = "2"
'        Call Gf_Common_DD(M_CN1, KeyCode)
'
'        If Len(Trim(txt_aplyspec_code.Text)) = txt_aplyspec_code.MaxLength Then
'            txt_aplystdspec.Text = Gf_ComnNameFind(M_CN1, "G0018", txt_aplyspec_code.Text, 2)
'            Exit Sub
'        Else
'            txt_aplystdspec.Text = ""
'        End If
'    End If
'End Sub

Private Sub txt_cur_inv_code_DblClick()

    Call txt_cur_inv_code_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_stdspec_chg_ref_DblClick()

    Call txt_stdspec_chg_ref_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_stdspec_chg_ref_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_stdspec_chg_ref

        Call Gf_StdSPEC_DD(M_CN1, KeyCode)

        Exit Sub

    End If
End Sub

Private Sub txt_cur_inv_code_Change()

    If Len(Trim(txt_cur_inv_code.Text)) = txt_cur_inv_code.MaxLength Then
          txt_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", txt_cur_inv_code.Text, 2)
          Exit Sub
    Else
          txt_cur_inv.Text = ""
    End If
    
End Sub

Private Sub txt_cur_inv_code_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "C0013"

        DD.rControl.Add Item:=txt_cur_inv_code
        DD.rControl.Add Item:=txt_cur_inv
        

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
       
        If Len(Trim(txt_cur_inv_code.Text)) = txt_cur_inv_code.MaxLength Then
            txt_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", txt_cur_inv_code.Text, 2)
            Exit Sub
        Else
            txt_cur_inv.Text = ""
        End If
    End If
    
End Sub

