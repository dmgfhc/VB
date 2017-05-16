VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AQE2000C 
   Caption         =   "质量异议及事件界面_AQE2000C"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   1890
   ClientWidth     =   12705
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8100
   ScaleWidth      =   12705
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_dept_name 
      Height          =   315
      Left            =   7290
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   135
      Width           =   1680
   End
   Begin VB.TextBox txt_dept 
      BackColor       =   &H00FFFFFF&
      Height          =   310
      Left            =   6660
      MaxLength       =   3
      TabIndex        =   10
      Tag             =   "dept"
      Top             =   135
      Width           =   600
   End
   Begin VB.TextBox txt_EMP_NAME 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   3495
      TabIndex        =   9
      Top             =   120
      Width           =   1575
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   30
      Left            =   0
      TabIndex        =   8
      Top             =   8070
      Width           =   12705
      _ExtentX        =   22410
      _ExtentY        =   53
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   4455
      Begin VB.OptionButton opt_KND_Q 
         BackColor       =   &H00E0E0E0&
         Caption         =   "操作人员违章/质量事件界面"
         Height          =   285
         Left            =   1800
         TabIndex        =   5
         Top             =   120
         Width           =   2625
      End
      Begin VB.OptionButton opt_KND_C 
         BackColor       =   &H00E0E0E0&
         Caption         =   "质量异议界面"
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Value           =   -1  'True
         Width           =   1425
      End
   End
   Begin VB.TextBox txt_KND 
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
      Left            =   9600
      MaxLength       =   14
      TabIndex        =   2
      Tag             =   "CD_MANA_NO"
      Top             =   720
      Visible         =   0   'False
      Width           =   930
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8085
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   15105
      _Version        =   393216
      _ExtentX        =   26644
      _ExtentY        =   14261
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   29
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQE2000C.frx":0000
   End
   Begin VB.TextBox txt_EMP_CD 
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
      Left            =   1440
      MaxLength       =   14
      TabIndex        =   0
      Tag             =   "CD_MANA_NO"
      Top             =   120
      Width           =   930
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   120
      Top             =   120
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      Caption         =   "责任人员ID"
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
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel12 
      Height          =   315
      Index           =   2
      Left            =   9060
      Top             =   135
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "发生日期"
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
   Begin InDate.UDate dtp_ISU_DATE_FROM 
      Height          =   300
      Left            =   10380
      TabIndex        =   6
      Tag             =   "订单确认日期"
      Top             =   135
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      BackColor       =   16777215
   End
   Begin InDate.UDate dtp_ISU_DATE_TO 
      Height          =   300
      Left            =   11820
      TabIndex        =   7
      Tag             =   "订单确认日期"
      Top             =   135
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
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
      Left            =   2445
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "人员姓名"
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
   Begin InDate.ULabel ULabel7 
      Height          =   300
      Left            =   5295
      Top             =   135
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      Caption         =   "责任单位"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
End
Attribute VB_Name = "AQE2000C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       质量异议及事件界面_AQE2000C
'-- Sub_System Name
'-- Program Name
'-- Program ID        AQE2000C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Sun Bin
'-- Coder
'-- Date              2006.9.6
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
Dim bCopy As Boolean
Dim lCopyRow As Long                'Copy Row


Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
           Call Gp_Ms_Collection(txt_EMP_CD, "P", " ", " ", " ", " ", "r", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(dtp_ISU_DATE_FROM, "P", " ", " ", " ", " ", "r", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(dtp_ISU_DATE_TO, "P", " ", " ", " ", " ", "r", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_KND, "P", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_EMP_NAME, " ", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'     Call Gp_Ms_Collection(dtp_PROD_DATE_TO, " ", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'         Call Gp_Ms_Collection(txt_Design_STS, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_dept, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_dept_name, " ", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
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
    Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 19, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AQE2000C.P_REFER", Key:="P-R"
    sc1.Add Item:="AQE2000C.P_ONEROW", Key:="P-O"
    sc1.Add Item:="AQE2000C.P_MODIFY", Key:="P-M"
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
    
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '--------------------------------------------------- Form_Activate --------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub Form_Activate()
         
        Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
    End Sub
    
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '--------------------------------------------------- Form_Load --------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub Form_Load()
    
        Screen.MousePointer = vbHourglass
        
        sAuthority = Gf_Pgm_Authority(Me.Name)
           
        Call Form_Define
    
        Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
        
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
        Call GP_ROW_BACKCOLOR(ss1)
        Call Gf_Sp_Cls(Proc_Sc("Sc"))
        Call Gp_Sp_HdColColor2(Proc_Sc("Sc")("Spread"), 2)
        Call Gp_Sp_HdColColor2(Proc_Sc("Sc")("Spread"), 3)
        Call Gp_Sp_HdColColor2(Proc_Sc("Sc")("Spread"), 5)
        Call Gp_Sp_HdColColor2(Proc_Sc("Sc")("Spread"), 10)
        Call Gp_Sp_HdColColor2(Proc_Sc("Sc")("Spread"), 17)
        Call Gp_Sp_HdColColor2(Proc_Sc("Sc")("Spread"), 19)
      
        txt_KND.Text = "C"
        dtp_ISU_DATE_FROM.Text = Date
        dtp_ISU_DATE_TO.Text = Date
            
        bCopy = False
            
        Screen.MousePointer = vbDefault
    
    End Sub
    
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '--------------------------------------------------- Form_QueryUnload ------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
        If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
            Cancel = 1
            Exit Sub
        End If
    
        Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
        
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
    
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '--------------------------------------------------- Form_KeyPress --------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub Form_KeyPress(KeyAscii As Integer)
        
        If KeyAscii = KEY_RETURN Then
            KeyAscii = 0
            SendKeys "{TAB}"
        End If

End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Form_KeyUp --------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)


    Select Case Me.ActiveControl.Name
              Case "txt_EMP_CD"
                 If KeyCode = vbKeyF4 Then
                 DD.sWitch = "MS"
                 DD.rControl.Add Item:=txt_EMP_CD
                 DD.rControl.Add Item:=txt_EMP_NAME
                 DD.nameType = "2"
                 Call Gf_EmpID_DD(M_CN1, KeyCode)
                End If
              Case "ss1"
               Call SP1_KeyUp(KeyCode, Shift)
    End Select
    
    
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Form_Ref ------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Public Sub Form_Ref()
    
On Error GoTo Refer_Err
    
    If dtp_ISU_DATE_FROM.RawData = "" Then
       dtp_ISU_DATE_FROM.Text = Date
    End If
    
    If dtp_ISU_DATE_TO.RawData = "" Then
       dtp_ISU_DATE_TO.Text = Date
    End If
    
     Call Master_To_Spread
     
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
            
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
    Call Txt_Knd_Value
    
    Exit Sub

Refer_Err:
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Form_Ins ------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Public Sub Form_Ins()

    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    Call Back_Colour
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 23)
    
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Form_Pro ------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Public Sub Form_Pro()

    Dim sMin As String
    Dim sMax As String
    
    Call Master_To_Spread
'    Call SSmesage_Error
    Call STS_SET
    If Gf_Sp_Process(M_CN1, Proc_Sc("Sc"), Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
    End If

    
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Form_Del ------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Public Sub Form_Del()

    If Not Gf_Ms_AllDel(M_CN1, Proc_Sc("Sc"), Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    End If

End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Form_Cls ------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("pControl"), False)
        pControl(1).SetFocus
        bCopy = False
        dtp_ISU_DATE_FROM.Text = ""
        dtp_ISU_DATE_TO.Text = ""
        txt_EMP_CD.Text = ""
        txt_EMP_NAME = ""
        
    End If

End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Form_Exc ------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Form_Exit ------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Public Sub Form_Exit()
    Unload Me
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- ss1_BlockSelected ------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- ss1_Change ------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub ss1_Change(ByVal Col As Long, ByVal Row As Long)

    If Gf_Sc_Authority(sAuthority, "U") Then

        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), 0)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 27)

    End If
    With ss1
      .Col = 3
      If .Text = "" Then
      .Col = 4
      .Text = ""
      End If
      .Col = 5
       If .Text = "" Then
      .Col = 6
      .Text = ""
      End If
      .Col = 10
      If .Text = "" Then
      .Col = 11
      .Text = ""
'      .Col = 12
'      .Text = ""
      End If
      .Col = 17
      If .Text = "" Then
      .Col = 18
      .Text = ""
      End If
      .Col = 19
      If .Text = "" Then
      .Col = 20
      .Text = ""
      End If
    End With
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- ss1_Click ------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- ss1_EditMode ------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), 2)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 27)
    End If
    
End Sub

''---------------------------------------------------------------------------------------------------------------------------------------------
''--------------------------------------------------- ss1_LeaveRow ------------------------------------------------------------------------
''---------------------------------------------------------------------------------------------------------------------------------------------
'Private Sub ss1_LeaveRow(ByVal Row As Long, ByVal RowWasLast As Boolean, ByVal RowChanged As Boolean, ByVal AllCellsHaveData As Boolean, ByVal NewRow As Long, ByVal NewRowIsLast As Long, Cancel As Boolean)
'  Spread_to_Master  Call (ss1, NewRow)
'End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- ss1_LostFocus ------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Spread_ColumnsSort ------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1
    
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Spread_Forzens_Setting ------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Public Sub Spread_Forzens_Setting()

    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol
    
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Spread_Forzens_Cancel ------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Public Sub Spread_Forzens_Cancel()

    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = 0
    
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- ss1_RightClick ------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub


'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Spread_Can ------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Public Sub Spread_Can()

    Call GP_SELECT_ROW(ss1, ss1.Row)
    Call GP_ROW_CANCEL(Proc_Sc("SC"))
    Call Gp_Ms_ControlLock(Mc1("pControl"), True)
      
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Spread_Del ------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Public Sub Spread_Del()
    
     Call GP_SET_CELL_VALUE(ss1, ss1.Row, 0, "Delete")

End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Spread_Cpy ------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Public Sub Spread_Cpy()

    lCopyRow = ss1.ActiveRow

End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Spread_Pst ------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Public Sub Spread_Pst()

    Call GP_ROW_PASTE(Proc_Sc("Sc"), lCopyRow)
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 23)
    
    bCopy = True

End Sub
Private Sub opt_KND_C_Click()

    If opt_KND_C.Value = True Then
        txt_KND.Text = "C"
       Call Gp_Sp_ColHidden(ss1, 5, False)
       Call Gp_Sp_ColHidden(ss1, 6, False)
       Call Gp_Sp_HdChange2
       Call Form_Ref
    End If

End Sub

Private Sub opt_KND_Q_Click()
    If opt_KND_Q.Value = True Then
        txt_KND.Text = "Q"
       Call Gp_Sp_ColHidden(ss1, 5, True)
       Call Gp_Sp_ColHidden(ss1, 6, True)
       Call Gp_Sp_ColHidden(ss1, 9, True)
       Call Gp_Sp_HdChange1
       Call Form_Ref
    End If
End Sub
Public Sub Gp_Sp_HdColColor2(sPname As Variant, iCol As Variant)

    With sPname
    
        .Row = 1: .Row2 = 1
        .Col = iCol: .Col2 = iCol
        
        .BlockMode = True
        
        .CellType = SS_CELL_TYPE_STATIC_TEXT
        .TypeHAlign = SS_CELL_H_ALIGN_CENTER
        .TypeVAlign = SS_CELL_V_ALIGN_CENTER
        .TypeTextWordWrap = True
        
        .BackColor = &HE1E4CD
        .ForeColor = BLUE
        
        .BlockMode = False
        
    End With
    
End Sub
Private Sub Gp_Sp_HdChange1()
With ss1
      .Row = 0
      .Col = 3
      .Text = "操作类型"
      .Col = 10
      .Text = "发生原因"
      
End With

End Sub
Private Sub Gp_Sp_HdChange2()
With ss1
      .Row = 0
      .Col = 3
      .Text = "异议类型"
      .Col = 10
      .Text = "异议原因"
      
End With

End Sub
Private Sub SP1_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Dim sTemp_Code As String

    If ss1.MaxRows < 1 Then Exit Sub
    
    If KeyCode <> vbKeyF4 Then
        Exit Sub
    End If

    Select Case ss1.ActiveCol
    
        Case 3
        
            If KeyCode = vbKeyF4 Then
                If txt_KND.Text = "C" Then
            
                   Set DD.sPname = Me.ss1
                
                       DD.sWitch = "SP"
                       DD.sKey = "Q0065"
                       DD.rControl.Add Item:=3
                       DD.rControl.Add Item:=4
                
                       DD.nameType = "1"
                
                       Call Gf_Common_DD(M_CN1, KeyCode)
                
                ElseIf txt_KND.Text = "Q" Then
                   Set DD.sPname = Me.ss1
                
                       DD.sWitch = "SP"
                       DD.sKey = "Q0066"
                       DD.rControl.Add Item:=3
                       DD.rControl.Add Item:=4
                
                       DD.nameType = "1"
                
                       Call Gf_Common_DD(M_CN1, KeyCode)
                End If

            End If
            
        Case 10
            If KeyCode = vbKeyF4 Then
            
                   Set DD.sPname = Me.ss1
                
                       DD.sWitch = "SP"
                       DD.sKey = "Q0067"
                       DD.rControl.Add Item:=10
                       DD.rControl.Add Item:=11
                
                       DD.nameType = "2"
                
                       Call Gf_Common_DD(M_CN1, KeyCode)

            End If
        Case 17
            If KeyCode = vbKeyF4 Then

                   Set DD.sPname = Me.ss1

                       DD.sWitch = "SP"
                       DD.sKey = "Z0002"
                       DD.rControl.Add Item:=17
                       DD.rControl.Add Item:=18
                       
                       DD.nameType = "2"

                       Call Gf_Common_DD(M_CN1, KeyCode)

            End If
        Case 5
            If KeyCode = vbKeyF4 Then

                Set DD.sPname = Me.ss1

                DD.sWitch = "SP"
                
                DD.rControl.Add Item:=5
                DD.rControl.Add Item:=6

                DD.nameType = "1"

                Call Gf_Customer_DD(M_CN1, KeyCode)

            End If

        Case 19
'            Dim dep_name As String
'
'            With ss1
'                .Col = 18
'                .Row = .ActiveRow
'                dep_name = .Text
'            End With
            If KeyCode = vbKeyF4 Then

                Set DD.sPname = Me.ss1

                DD.sWitch = "SP"
                DD.rControl.Add Item:=19
                DD.rControl.Add Item:=20
                DD.nameType = "2"

               Call Code_Name(M_CN1, KeyCode)
            
            End If

    End Select

End Sub
Public Sub Txt_Knd_Value()
    If opt_KND_C.Value = True Then
        txt_KND.Text = "C"
   ElseIf opt_KND_Q.Value = True Then
        txt_KND.Text = "Q"
   End If
End Sub

Private Sub Code_Name(Conn As ADODB.Connection, KeyCode As Integer)
    Dim sOld_Code, sNew_Code  As String
    Dim sOld_Name, sNew_Name  As String

    DD.DataDicType = "EMP"      'Program ID
    DD.DicRefType = "C"         'Active Form DataDic Call

        DD.sPname.Col = DD.rControl.Item(1)
        sOld_Code = DD.sPname.Text

        DD.sQuery = "            SELECT EMP_ID ""人员 ID"", EMP_NAME ""人员名称"" FROM  NISCO.ZP_EMPLOYEE "
'        DD.sWhere = "             WHERE DESCRIPTION  LIKE '" & Trim(DD.sKey) & "' "

    If Gf_DD_Display(Conn, DD.sQuery + DD.sWhere, False) Then

        If DD.sWitch = "SP" Then

            DD.sPname.Col = DD.rControl.Item(1)
            sNew_Code = DD.sPname.Text

            If DD.rControl.COUNT > 1 Then
                DD.sPname.Col = DD.rControl.Item(2)
                sNew_Name = DD.sPname.Text
            End If

            DD.sPname.TabStop = True
            DD.sPname.SetFocus
            DD.sPname.SetActiveCell DD.rControl.Item(1), DD.sPname.ActiveRow
            DD.sPname.Action = SS_ACTION_ACTIVE_CELL
            DD.sPname.EditMode = True
            DD.sPname.TabStop = False

            If DD.sSelect Then
                If sOld_Code <> sNew_Code Then Call Gp_Sp_UpdateMake(DD.sPname, False)
            End If
        End If

    End If

    DD.sWitch = ""
    DD.sSelect = False

    Set DD.sPname = Nothing
    Set DD.rControl = Nothing


End Sub
Private Sub STS_SET()
Dim REASON_CD_STS As String
Dim RSLT_STS As String

 With ss1
 .Col = 8
 .Row = .ActiveRow
 REASON_CD_STS = .Text
 .Col = 12
 RSLT_STS = .Text
 .Col = 2
 If REASON_CD_STS = "" And RSLT_STS = "" Then
    .Text = "A"
 ElseIf REASON_CD_STS <> "" Or RSLT_STS <> "" And .Text = "A" Then
    .Text = "B"
Else: Exit Sub
End If

 End With
End Sub
Private Sub Master_To_Spread()
With ss1
    .Col = 29
    .Row = .ActiveRow
    .Text = txt_KND.Text

End With

End Sub
'Private Sub SSmesage_Error()
'Dim REASON_CD_STS As String
'Dim RSLT_STS As String
'
' With ss1
' .Col = 8
' .Row = .ActiveRow
' REASON_CD_STS = .Text
' .Col = 12
' RSLT_STS = .Text
'.Col = 2
'If .Text = "A" Then
'  If REASON_CD_STS = "" Then
'  Call MsgBox("事件经过必需输入！", vbOKOnly, "系统提示信息")
'  ElseIf REASON_CD_STS <> "" And RSLT_STS = "" Then
'  Call MsgBox("原因详细必需输入！", vbOKOnly, "系统提示信息")
'  End If
'End If
'End With
'End Sub

Private Sub txt_emp_cd_Change()
Dim ID_CD As String
Dim sQuery As String
ID_CD = txt_EMP_CD.Text

If txt_EMP_CD.Text = "" Then
   txt_EMP_NAME.Text = ""
End If

If txt_EMP_CD.Text <> "" And Len(txt_EMP_CD) = 7 Then
   sQuery = "SELECT EMP_NAME FROM NISCO.ZP_EMPLOYEE WHERE EMP_ID='" + ID_CD + "'"
   txt_EMP_NAME.Text = Gf_FloatFind(M_CN1, sQuery)
 End If
End Sub
Private Sub Back_Colour()
    With ss1
        .Row = .ActiveRow
'        .Col = 2
'        .BackColor = &H80000005
        .Col = 7
        .BackColor = &HFFFF&
        .Col = 19
        .BackColor = &HFFFF&
    End With
End Sub

Private Sub txt_dept_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "Z0002"
        DD.rControl.Add Item:=txt_dept
        DD.rControl.Add Item:=txt_dept_name
        
        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If
    

    If Len(Trim(txt_dept.Text)) = txt_dept.MaxLength Then
        txt_dept_name.Text = Gf_ComnNameFind(M_CN1, "Z0002", txt_dept.Text, 2)
    Else
        txt_dept_name.Text = ""
    End If

End Sub



