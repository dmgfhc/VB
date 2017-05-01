VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form UserID_Add 
   BackColor       =   &H00FFFAE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UserID_Add"
   ClientHeight    =   1590
   ClientLeft      =   8160
   ClientTop       =   6645
   ClientWidth     =   5880
   Icon            =   "UserID_Add.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   5880
   Begin VB.TextBox txt_emp_id 
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
      Left            =   2385
      MaxLength       =   7
      TabIndex        =   0
      Top             =   135
      Width           =   2265
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
      Left            =   2385
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   540
      Width           =   2265
   End
   Begin Threed.SSCommand cmd_Cancel 
      Height          =   420
      Left            =   2295
      TabIndex        =   3
      Tag             =   "取消"
      Top             =   1080
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
   Begin Threed.SSCommand cmd_OK 
      Height          =   420
      Left            =   945
      TabIndex        =   2
      Tag             =   "確認"
      Top             =   1080
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
   Begin Threed.SSCommand cmd_Del 
      Height          =   420
      Left            =   3645
      TabIndex        =   4
      Top             =   1080
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
      Caption         =   "删除"
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
      Left            =   1215
      TabIndex        =   6
      Top             =   180
      Width           =   1050
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
      Left            =   1215
      TabIndex        =   5
      Top             =   585
      Width           =   1050
   End
End
Attribute VB_Name = "UserID_Add"
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
'-- Program Name      Confirm Login
'-- Program ID        ConfirmLogin
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

Private Sub cmd_Del_Click()

    UserID_Del.Show 1
    
End Sub

Private Sub Cmd_Ok_Click()
    
On Error GoTo Ok_Error
    
    Dim sQuery As String
    Dim md5_pwd As String
    Dim user_pwd As String

    'ID Check
    sQuery = "SELECT EMP_NAME FROM ZP_EMPLOYEE WHERE EMP_ID = '" + Trim(txt_emp_id.Text) + "'"
    
    If Gf_CodeFind(M_CN1, sQuery) = "" Then
        Call Gp_MsgBoxDisplay("您输入的用户代码不存在！", "W")
        Exit Sub
    End If
    
    sQuery = "SELECT gf_md5_pwd('" + Trim(txt_password.Text) + "') FROM dual"
    md5_pwd = Gf_CodeFind(M_CN1, sQuery)

    'Password check
    sQuery = "SELECT PASSWORD FROM ZP_EMPLOYEE WHERE EMP_ID = '" + Trim(txt_emp_id.Text) + "'"
    user_pwd = Gf_CodeFind(M_CN1, sQuery)
    
    If md5_pwd <> user_pwd Then
        Call Gp_MsgBoxDisplay("密码错误！", "I")
        Exit Sub
    End If
    
    'Registry ID Insert
    SaveSetting "NISCO", "LOGIN", txt_emp_id.Text, 1
    
    Unload Me
    Exit Sub
    
Ok_Error:

    Call Gp_MsgBoxDisplay("新增用户代码失败...")
    
End Sub

Private Sub cmd_cancel_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()

    Me.KeyPreview = True
    'Me.BackColor = &HE0E0E0
    
    Call Gp_FormCenter(Me)
    
End Sub
