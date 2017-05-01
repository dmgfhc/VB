VERSION 5.00
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form AGF2033C 
   Caption         =   "支撑辊轴承(座)保养管理界面_AGF2033C"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11490
   ScaleWidth      =   19080
   WindowState     =   2  'Maximized
   Begin VB.ComboBox CBO_STATION 
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
      ItemData        =   "AGF2033C.frx":0000
      Left            =   13800
      List            =   "AGF2033C.frx":000A
      TabIndex        =   4
      Tag             =   "位置"
      Top             =   120
      Width           =   1305
   End
   Begin VB.ComboBox CBO_OW 
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
      ItemData        =   "AGF2033C.frx":001A
      Left            =   11070
      List            =   "AGF2033C.frx":0024
      TabIndex        =   3
      Tag             =   "部位"
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox CBO_BEARING_ID 
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
      ItemData        =   "AGF2033C.frx":003C
      Left            =   1395
      List            =   "AGF2033C.frx":003E
      TabIndex        =   0
      Tag             =   "轴承号"
      Top             =   120
      Width           =   1365
   End
   Begin VB.OptionButton OPT_CH 
      BackColor       =   &H00E0E0E0&
      Caption         =   "轴承号"
      Height          =   315
      Left            =   2895
      TabIndex        =   8
      Top             =   120
      Value           =   -1  'True
      Width           =   930
   End
   Begin VB.OptionButton OPT_DISCH 
      BackColor       =   &H00E0E0E0&
      Caption         =   "轴承座号"
      Height          =   315
      Left            =   3855
      TabIndex        =   7
      Top             =   120
      Width           =   1110
   End
   Begin VB.TextBox TXT_CH_CD 
      Height          =   315
      Left            =   4815
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   315
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8535
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   14985
      _Version        =   393216
      _ExtentX        =   26432
      _ExtentY        =   15055
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
      MaxCols         =   21
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AGF2033C.frx":0040
   End
   Begin InDate.ULabel ULabel16 
      Height          =   315
      Left            =   120
      Top             =   120
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "轴承号"
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
      Left            =   5175
      Top             =   120
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "处理日期"
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
   Begin InDate.UDate SDT_DATE_FROM 
      Height          =   315
      Left            =   6465
      TabIndex        =   1
      Tag             =   "起始日期"
      Top             =   120
      Width           =   1470
      _ExtentX        =   2593
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
   Begin InDate.UDate SDT_DATE_TO 
      Height          =   315
      Left            =   8145
      TabIndex        =   2
      Tag             =   "起始日期"
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
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   9795
      Top             =   120
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "部位"
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
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   12525
      Top             =   120
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "位置"
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "～"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   7950
      TabIndex        =   9
      Top             =   150
      Width           =   255
   End
End
Attribute VB_Name = "AGF2033C"
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
'-- Program Name      支撑辊轴承(座)保养管理界面
'-- Program ID        AGF2033C
'-- Document No       Q-00-0010(Specification)
'-- Designer          ZHANG
'-- Coder             ZHANG
'-- Date              2009.7.24
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
'Dim Mc2 As New Collection           'Master Collection
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
         Call Gp_Ms_Collection(CBO_BEARING_ID, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(CBO_OW, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(CBO_STATION, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDT_DATE_FROM, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(SDT_DATE_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(TXT_CH_CD, "p", "n", " ", " ", " ", "a", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                   
     Mc1.Add Item:=pControl, Key:="pControl"
     Mc1.Add Item:=nControl, Key:="nControl"
     Mc1.Add Item:=mControl, Key:="mControl"
     Mc1.Add Item:=iControl, Key:="iControl"
     Mc1.Add Item:=rControl, Key:="rControl"
     Mc1.Add Item:=cControl, Key:="cControl"
     Mc1.Add Item:=aControl, Key:="aControl"
     Mc1.Add Item:=lControl, Key:="lControl"
               
     Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 2, "p", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 14, "p", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", "i", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", "i", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 21, "p", " ", " ", "i", "a", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
   
    
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AGF2033C.P_SREFER", Key:="P-R"
    sc1.Add Item:="AGF2033C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="AGF2033C.P_ONEROW", Key:="P-O"
    sc1.Add Item:=pColumn, Key:="pColumn"
    sc1.Add Item:=nColumn, Key:="nColumn"
    sc1.Add Item:=aColumn, Key:="aColumn"
    sc1.Add Item:=mColumn, Key:="mColumn"
    sc1.Add Item:=iColumn, Key:="iColumn"
    sc1.Add Item:=lColumn, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    Call Gp_Sp_ColHidden(ss1, 2, True)

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").Row = 0
    sc1.Item("Spread").Text = "◎"
    

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
    
'   Call Gp_Sp_ColHidden(ss1, 11, True)
'   Call Gp_Sp_ColHidden(ss1, 12, True)
'   Call Gp_Sp_ColHidden(ss1, 13, True)
'
    
    TXT_CH_CD = "1"
    If OPT_CH.Value = True Then
       OPT_CH.ForeColor = &HFF&
       OPT_DISCH.ForeColor = &H808080
    Else
       OPT_CH.ForeColor = &H808080
    End If
    
    ULabel16.Caption = "轴承号"
    sQuery_load = "SELECT BEARING_ID FROM GP_BEARING WHERE  BEARING_ID LIKE 'B2%' AND STATUS <> 'DL' ORDER BY BEARING_ID "
    Call Gf_ComboAdd(M_CN1, CBO_BEARING_ID, sQuery_load)
    
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

Public Sub Spread_Can()
Dim sCid As String
    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
    ss1.Row = ss1.ActiveRow
    ss1.Col = 2
    sCid = ss1.Text
    If sCid <> "" Then
       ss1.Col = 1
       ss1.Text = sCid
    End If
      
End Sub
Public Sub Form_Cls()
    
    If Gf_Sp_Cls(sc1) Then
       Call Gp_Ms_Cls(Mc1("rControl"))
       Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
       Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        SDT_DATE_FROM.SetFocus
        SDT_DATE_FROM.Text = ""
        SDT_DATE_TO.Text = ""
    End If
End Sub

Public Sub Form_Ref()
Dim I As Integer
Dim sCid As String

On Error GoTo Refer_Err

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        ss1.OperationMode = OperationModeNormal
        For I = 1 To ss1.MaxRows
            ss1.Col = 2
            ss1.Row = I
            sCid = ss1.Text
            If sCid <> "" Then
               ss1.Col = 1
               ss1.Text = sCid
            End If
        Next I
    End If
Refer_Err:
End Sub

Public Sub Form_Pro()
     
If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
  Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
  Call Form_Ref
End If
     
End Sub


Public Sub Form_Ins()
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    ss1.Col = 16
    ss1.Row = ss1.ActiveRow
    ss1.Text = sUserID
    
    Call Gp_Sp_ColLock(ss1, 1, False)
    
    ss1.Row = ss1.ActiveRow
    ss1.Col = 1
    ss1.BackColor = &HC0FFFF
    
    If OPT_CH.Value = True Then
        Call Pf_ComboAdd(M_CN1, ss1, 1, "SELECT BEARING_ID  FROM GP_BEARING WHERE BEARING_ID LIKE 'B2%' AND STATUS <> 'DL' ORDER BY BEARING_ID ")
    ElseIf OPT_DISCH.Value = True Then
        Call Pf_ComboAdd(M_CN1, ss1, 1, "SELECT CHOCK_ID  FROM GP_CHOCK WHERE CHOCK_ID LIKE 'C2%' AND STATUS <> 'DL' ORDER BY CHOCK_ID ")
    End If
    
End Sub

Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    Proc_Sc("Sc").Item("Spread").OperationMode = OperationModeNormal
    
End Sub
Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    ss1.Col = 16
    ss1.Row = ss1.ActiveRow
    ss1.Text = sUserID
    
    Call Gp_Sp_ColLock(ss1, 1, False)
    
    ss1.Row = ss1.ActiveRow
    ss1.Col = 1
    ss1.BackColor = &HC0FFFF
    
    If OPT_CH.Value = True Then
        Call Pf_ComboAdd(M_CN1, ss1, 1, "SELECT BEARING_ID  FROM GP_BEARING WHERE STATUS <> 'DL' ORDER BY BEARING_ID ")
    ElseIf OPT_DISCH.Value = True Then
        Call Pf_ComboAdd(M_CN1, ss1, 1, "SELECT CHOCK_ID  FROM GP_CHOCK WHERE STATUS <> 'DL' ORDER BY CHOCK_ID ")
    End If

    ss1.OperationMode = OperationModeNormal

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
Private Sub OPT_CH_Click()
    Dim iRow As Integer
    Dim sTemp As String
    
    Call Form_Cls
   
    ss1.Row = 0: ss1.Col = 1
    ss1.Text = "轴承号"
    
    ULabel16.Caption = "轴承号"
    sQuery_load = "SELECT BEARING_ID FROM GP_BEARING  WHERE BEARING_ID LIKE 'B2%' AND STATUS <> 'DL' ORDER BY BEARING_ID "
    Call Gf_ComboAdd(M_CN1, CBO_BEARING_ID, sQuery_load)

    If OPT_CH.Value = True Then
        OPT_CH.ForeColor = &HFF&
        OPT_DISCH.ForeColor = &H808080
    
     TXT_CH_CD = "1"
    Else
        OPT_CH.ForeColor = &H808080
    End If


End Sub


Private Sub OPT_DISCH_Click()
    Dim iRow As Integer
    Dim sTemp As String

    Call Form_Cls

    ss1.Row = 0: ss1.Col = 1
    ss1.Text = "轴承座号"
    ULabel16.Caption = "轴承座号"
    sQuery_load = "SELECT CHOCK_ID FROM GP_CHOCK  WHERE CHOCK_ID LIKE 'C2%' AND STATUS <> 'DL' ORDER BY CHOCK_ID "
    Call Gf_ComboAdd(M_CN1, CBO_BEARING_ID, sQuery_load)

    
    If OPT_DISCH.Value = True Then
        OPT_DISCH.ForeColor = &HFF&
        OPT_CH.ForeColor = &H808080

        TXT_CH_CD = "2"
    Else
        OPT_DISCH.ForeColor = &H808080
    End If
End Sub



Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
Dim sCid As String
    
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub
Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
Dim sCid As String
If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        ss1.Col = 0
        ss1.Row = ss1.ActiveRow
        If ss1.Text = "Update" Then
            ss1.Col = 18
            ss1.Text = sUserID
        End If
        
        If Col = 1 Then
           ss1.Col = 1
           sCid = ss1.Text
           ss1.Col = 2
           ss1.Text = sCid
        End If
    End If
    
End Sub

Private Sub ss1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim sCid As String
If ss1.ActiveCol = 1 Then
   ss1.Row = ss1.ActiveRow
   ss1.Col = ss1.ActiveCol
   sCid = Trim(ss1.Text)
   If Len(sCid) = 7 Then
      ss1.Col = 2
      ss1.Text = sCid
   End If
End If
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


Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)

    If Row = 0 Then Exit Sub
    ss1.Row = Row
    ss1.Col = Col
     
'  If ss1.Lock = False Then
'        If ss1.Col = 3 Then
'
'         ss1.Text = Format(Now, "YYYY-MM-DD HH:MM:SS")
'
'         ss1.Col = 0
'         Select Case Trim(ss1.Text)
'                Case "Input", "Update", "Delete"
'                Case Else
'                ss1.Text = "Update"
'         End Select
'       End If
'    End If

    If ss1.Col = 14 Then

        ss1.Text = Format(Now, "YYYY-MM-DD")
        
    End If
    
    If ss1.Col = 15 Then
                   
        ss1.Text = Format(Now, "YYYY-MM-DD")

    End If

    
End Sub

Public Function Pf_ComboAdd(Conn As ADODB.Connection, ss As vaSpread, Col As Integer, sQuery As String, Optional ClsChk As Boolean = True) As Boolean

On Error GoTo ComboAdd_Error

    Dim AdoRs As ADODB.Recordset
    Dim sList As String
    
    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Pf_ComboAdd = False: Exit Function
    End If
    
'    If ClsChk Then
'        Cbo.Clear
'    End If
    
    Set AdoRs = New ADODB.Recordset

    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
        While Not AdoRs.EOF
            
            If AdoRs.Fields(0) <> vbNull Then
                sList = sList & AdoRs.Fields(0) & vbTab
                'Cbo.AddItem AdoRs.Fields(0)
            End If
            AdoRs.MoveNext
            
        Wend
        Pf_ComboAdd = True
    Else
        Pf_ComboAdd = False
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    ss.Col = Col
    ss.TypeComboBoxList = sList
    
    Exit Function

ComboAdd_Error:

    Set AdoRs = Nothing
    Pf_ComboAdd = False

End Function
