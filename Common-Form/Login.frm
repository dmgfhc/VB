VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form Login 
   BackColor       =   &H00FFFAE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   4290
   ClientLeft      =   4800
   ClientTop       =   1770
   ClientWidth     =   5865
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   5865
   Begin VB.ComboBox cbo_emp_id 
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
      Left            =   3105
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2970
      Width           =   1680
   End
   Begin VB.TextBox txt_password 
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
      IMEMode         =   3  'DISABLE
      Left            =   3105
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3375
      Width           =   1680
   End
   Begin Threed.SSCommand Cmd_OK 
      Height          =   420
      Left            =   1800
      TabIndex        =   2
      Tag             =   "確認"
      Top             =   3780
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      _Version        =   196609
      Font3D          =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "确认"
   End
   Begin Threed.SSCommand Cmd_Cancel 
      Height          =   420
      Left            =   3150
      TabIndex        =   3
      Tag             =   "取消"
      Top             =   3780
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      _Version        =   196609
      Font3D          =   2
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
   Begin Threed.SSCommand Cmd_Change 
      Height          =   465
      Left            =   4860
      TabIndex        =   5
      Tag             =   "变更"
      Top             =   3240
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   820
      _Version        =   196609
      Font3D          =   2
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Login.frx":0CCA
   End
   Begin Threed.SSCommand Cmd_Setting 
      Height          =   420
      Left            =   4500
      TabIndex        =   4
      Tag             =   "設定"
      Top             =   3780
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      _Version        =   196609
      Font3D          =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "设定"
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "密码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1935
      TabIndex        =   7
      Top             =   3420
      Width           =   1050
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "用户代码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1935
      TabIndex        =   6
      Top             =   3015
      Width           =   1050
   End
   Begin VB.Image Image1 
      Height          =   2925
      Left            =   0
      Picture         =   "Login.frx":0EAF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5880
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name
'-- Sub_System Name
'-- Program Name      Login
'-- Program ID        Login
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
'-- Date              2003.6.9
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------

Public PassCnt As Integer           'PassWord Input Count

Private Sub Cmd_Ok_Click()

    Dim sQuery As String
    Dim md5_pwd As String
    Dim user_pwd As String
    
    'ID Check
    sQuery = "SELECT EMP_NAME FROM ZP_EMPLOYEE WHERE EMP_ID = '" + Trim(cbo_emp_id.Text) + "'"
    sUserName = Gf_CodeFind(M_CN1, sQuery)
    
    If sUserName = "" Then
        Call Gp_MsgBoxDisplay("您输入的用户代码不存在", "I")
        M_CN1.Close
        Set M_CN1 = Nothing
        Exit Sub
    End If
    
    sQuery = "SELECT gf_md5_pwd( '" + Trim(txt_password.Text) + "') FROM dual"
    md5_pwd = Gf_CodeFind(M_CN1, sQuery)
    
    'Password check
    sQuery = "SELECT PASSWORD FROM ZP_EMPLOYEE WHERE EMP_ID = '" + Trim(cbo_emp_id.Text) + "'"
    user_pwd = Gf_CodeFind(M_CN1, sQuery)
    
    If md5_pwd <> user_pwd Then
        
        PassCnt = PassCnt + 1
        
        If PassCnt > 2 Then
            Call Gp_MsgBoxDisplay("密码错误三次" + vbCrLf + "请退出系统..")
            PassCheck = False
            M_CN1.Close
            Set M_CN1 = Nothing
            Unload Me
            Exit Sub
        End If
        
        Call Gp_MsgBoxDisplay("密码错误", "I")
        M_CN1.Close
        Set M_CN1 = Nothing
    Else
    
        sUserID = Trim(cbo_emp_id.Text)
        
        'Current User Registry Setting
        SaveSetting "NISCO", "AUTHORITY", "sUserID", sUserID
        SaveSetting "NISCO", "AUTHORITY", "sUsername", sUserName

        Call Gp_DateSetting
        PassCheck = True
        M_CN1.Close
        Set M_CN1 = Nothing
        Unload Me
        
    End If
    
End Sub

Private Sub cmd_cancel_Click()
    
    PassCheck = False
    Unload Me
    
End Sub

Private Sub cmd_change_Click()

    Dim sQuery As String
    Dim md5_pwd As String
    Dim user_pwd As String
    
    'ID Check
    sQuery = "SELECT EMP_NAME FROM ZP_EMPLOYEE WHERE EMP_ID = '" + Trim(cbo_emp_id.Text) + "'"
    sUserName = Gf_CodeFind(M_CN1, sQuery)
    
    If sUserName = "" Then
        Call Gp_MsgBoxDisplay("您输入的用户代码不存在", "I")
        Exit Sub
    End If
    
    sQuery = "SELECT gf_md5_pwd('" + Trim(txt_password.Text) + "') FROM dual"
    md5_pwd = Gf_CodeFind(M_CN1, sQuery)

    'Password check
    sQuery = "SELECT PASSWORD FROM ZP_EMPLOYEE WHERE EMP_ID = '" + Trim(cbo_emp_id.Text) + "'"
    user_pwd = Gf_CodeFind(M_CN1, sQuery)
    
    If md5_pwd <> user_pwd Then
        
        Call Gp_MsgBoxDisplay("密码错误", "I")
        PassCnt = PassCnt + 1
        Exit Sub
    Else
        sUserID = Trim(cbo_emp_id.Text)
    End If

    ConfirmLogin.Show 1
    
End Sub

Private Sub Cmd_Setting_Click()

    Dim Reg_id As Variant
    Dim intSettings As Integer
    
    UserID_Add.Show 1
    
    'Registry ID Setting
    Reg_id = GetAllSettings("NISCO", "LOGIN")
    
    cbo_emp_id.Clear
    
    For intSettings = LBound(Reg_id, 1) To UBound(Reg_id, 1)
        If Reg_id(intSettings, 0) <> "SAMPLE" Then
            cbo_emp_id.AddItem Reg_id(intSettings, 0)
        End If
    Next intSettings
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()
    
    'Dim WshShell As Object
    Dim Reg_id As Variant
    Dim intSettings As Integer
    
    PassCheck = False
    
    Me.KeyPreview = True
    'Me.BackColor = &HE0E0E0
    
    'CapsLock On
    If GetKeyState(KEY_CAPITAL) <> 1 Then
        Call keybd_event(KEY_CAPITAL, 0, 0, 0)
    End If
    
    'Set WshShell = CreateObject("WScript.Shell")
    'WshShell.SendKeys "{CAPSLOCK}"
    
    Call Gp_FormCenter(Me)
    
    'Registry ID Setting
    SaveSetting "NISCO", "LOGIN", "SAMPLE", 1
    Reg_id = GetAllSettings("NISCO", "LOGIN")
    
    cbo_emp_id.Clear
    
    For intSettings = LBound(Reg_id, 1) To UBound(Reg_id, 1)
        If Reg_id(intSettings, 0) <> "SAMPLE" Then
            cbo_emp_id.AddItem Reg_id(intSettings, 0)
        End If
    Next intSettings
    
End Sub

