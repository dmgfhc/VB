VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form SCQKJBZY 
   Caption         =   "连铸生产情况及班注意事项"
   ClientHeight    =   6225
   ClientLeft      =   4050
   ClientTop       =   3345
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6225
   ScaleWidth      =   8355
   Begin VB.TextBox Txt_five 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   720
      MaxLength       =   100
      TabIndex        =   10
      Tag             =   "txt_five"
      Top             =   2325
      Width           =   7155
   End
   Begin VB.TextBox Txt_ten 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   720
      MaxLength       =   100
      TabIndex        =   9
      Tag             =   "txt_five"
      Top             =   4950
      Width           =   7155
   End
   Begin VB.TextBox txt_nine 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   720
      MaxLength       =   100
      TabIndex        =   8
      Tag             =   "作业人员"
      Top             =   4425
      Width           =   7155
   End
   Begin VB.TextBox Txt_eight 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   720
      MaxLength       =   100
      TabIndex        =   7
      Tag             =   "作业人员"
      Top             =   3900
      Width           =   7155
   End
   Begin VB.TextBox Txt_seven 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   720
      MaxLength       =   100
      TabIndex        =   6
      Tag             =   "作业人员"
      Top             =   3375
      Width           =   7155
   End
   Begin VB.TextBox txt_six 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   720
      MaxLength       =   100
      TabIndex        =   5
      Tag             =   "作业人员"
      Top             =   2850
      Width           =   7155
   End
   Begin VB.TextBox Txt_two 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   720
      MaxLength       =   100
      TabIndex        =   4
      Tag             =   "作业人员"
      Top             =   765
      Width           =   7155
   End
   Begin VB.TextBox Txt_one 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   720
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   240
      Width           =   7155
   End
   Begin VB.TextBox Txt_four 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   720
      MaxLength       =   100
      TabIndex        =   2
      Tag             =   "作业人员"
      Top             =   1815
      Width           =   7155
   End
   Begin VB.TextBox Txt_three 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   720
      MaxLength       =   100
      TabIndex        =   1
      Tag             =   "作业人员"
      Top             =   1290
      Width           =   7155
   End
   Begin VB.TextBox txt_m_shift 
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
      Left            =   1575
      MaxLength       =   1
      TabIndex        =   0
      Top             =   5565
      Visible         =   0   'False
      Width           =   1200
   End
   Begin InDate.ULabel ULabel5 
      Height          =   510
      Left            =   390
      Top             =   2325
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   900
      Caption         =   "5"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel6 
      Height          =   510
      Left            =   390
      Top             =   2850
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   900
      Caption         =   "6"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel7 
      Height          =   510
      Left            =   390
      Top             =   3375
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   900
      Caption         =   "7"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel8 
      Height          =   510
      Left            =   390
      Top             =   3900
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   900
      Caption         =   "8"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel9 
      Height          =   510
      Left            =   390
      Top             =   4425
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   900
      Caption         =   "9"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel10 
      Height          =   510
      Left            =   390
      Top             =   4950
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   900
      Caption         =   "10"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   510
      Left            =   390
      Top             =   240
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   900
      Caption         =   "1"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel2 
      Height          =   510
      Left            =   390
      Top             =   765
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   900
      Caption         =   "2"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel4 
      Height          =   510
      Left            =   390
      Top             =   1815
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   900
      Caption         =   "4"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel3 
      Height          =   510
      Left            =   390
      Top             =   1290
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   900
      Caption         =   "3"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.UDate txt_m_date 
      Height          =   315
      Left            =   15
      TabIndex        =   11
      Tag             =   "起始日期"
      Top             =   5565
      Visible         =   0   'False
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
   Begin Threed.SSCommand cmd_OK 
      Height          =   435
      Left            =   3390
      TabIndex        =   12
      Top             =   5640
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   767
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "确定"
   End
   Begin Threed.SSCommand cmd_Cancel 
      Height          =   435
      Left            =   4770
      TabIndex        =   13
      Top             =   5640
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   767
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "取消"
   End
End
Attribute VB_Name = "SCQKJBZY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       NISCO Production Management System
'-- Sub_System Name   Steel Making System
'-- Program Name      LF RESLT MODIFICATION
'-- Program ID        AFB2010C
'-- Document No
'-- Designer          H.M.G
'-- Coder             H.M.G
'-- Date              2003.7.23
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
Public sDateTime As String          'Active Form Authority Setting
Public sQuery_Rt As String          'Active Form Authority Setting

Dim pControl As New Collection      'Master Primary Key Collection
Dim nControl As New Collection      'Master Necessary Collection
Dim mControl As New Collection      'Master Maxlength check Collection
Dim iControl As New Collection      'Master Insert Collection
Dim rControl As New Collection      'Master Refer Collection
Dim cControl As New Collection      'Master Copy Collection
Dim aControl As New Collection      'Master -> Spread Collection
Dim lControl As New Collection      'Master Lock Collection

Dim Mc1 As New Collection           'Master Collection




Private Sub Form_Define()
       
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Master"              'form类型
       'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
  Call Gp_Ms_Collection(txt_m_date, "p", "n", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(txt_m_shift, "p", "n", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(Txt_one, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(Txt_two, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(Txt_three, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(Txt_four, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(Txt_five, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_six, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(Txt_seven, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(Txt_eight, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_nine, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(Txt_ten, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                   
           
    'MASTER Collection
     Mc1.Add Item:="AFH5012C.P_MODIFY", Key:="P-M"
     Mc1.Add Item:="AFH5012C.P_SREFER", Key:="P-R"
     Mc1.Add Item:=pControl, Key:="pControl"
     Mc1.Add Item:=nControl, Key:="nControl"
     Mc1.Add Item:=mControl, Key:="mControl"
     Mc1.Add Item:=iControl, Key:="iControl"
     Mc1.Add Item:=rControl, Key:="rControl"
     Mc1.Add Item:=cControl, Key:="cControl"
     Mc1.Add Item:=aControl, Key:="aControl"
     Mc1.Add Item:=lControl, Key:="lControl"
     
     Me.KeyPreview = True
     Me.BackColor = &HE0E0E0

End Sub
 
Private Sub Command2_Click()
   Call Form_Exit
End Sub

Private Sub Cmd_Cancel_Click()
  Unload Me
End Sub

Private Sub Cmd_Ok_Click()
    If txt_m_date = "" Then
       Call Gp_MsgBoxDisplay("日期必须输入", "", "错误提示")
 
       Unload Me
    ElseIf txt_m_shift = "" Then
       Call Gp_MsgBoxDisplay("班次必须输入", "", "错误提示")
   
       Unload Me
    End If
    
    Call Gf_Ms_Process(M_CN1, Mc1, sAuthority)
    
'     If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then
'       Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
'    End If
  '  Call AFH5010C.Form_Ref
  '  Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    Call Gp_FormCenter(Me)
    
    Call Form_Define
    
'    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    
   ' Call Gp_Ms_ControlLock(Mc1("lControl"), True)
    
   ' Call Gp_Ms_NeceColor(Mc1("nControl"))
    Screen.MousePointer = vbDefault
    
  If Mid(sAuthority, 3, 1) <> "1" Then
     cmd_ok.Enabled = False
  ElseIf Mid(sAuthority, 3, 1) = "1" Then
     cmd_ok.Enabled = True
  End If
End Sub
Public Sub Form_Ref()
  
    Call Gf_Ms_Refer(M_CN1, Mc1, Mc1("pControl"), Mc1("mControl"), False)
    
'      If Gf_Ms_Refer(M_CN1, Mc1, Mc1("pControl"), Mc1("mControl"), False) Then
'          Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'         ' Call Gp_Ms_ControlLock(Mc1("pControl"), False)
'      End If
      
End Sub
Public Sub Form_Exit()

    Unload Me
    
End Sub

Public Sub Form_Cls()

    Call Gp_Ms_Cls(Mc1("rControl"))
'    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
End Sub
 
Private Sub Form_Activate()

'    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set Mc1 = Nothing

'    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub
