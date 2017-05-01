VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form ConfirmLogin 
   BackColor       =   &H00FFFAE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "密码确认"
   ClientHeight    =   1590
   ClientLeft      =   2130
   ClientTop       =   6660
   ClientWidth     =   5880
   Icon            =   "ConfirmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   5880
   Begin VB.TextBox txt_confirm 
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
      Left            =   2655
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   540
      Width           =   2535
   End
   Begin Threed.SSCommand cmd_Cancel 
      Height          =   420
      Left            =   2970
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
      Left            =   1620
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
   Begin VB.TextBox txt_newpassword 
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
      Left            =   2655
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   135
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "新密码"
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
      Left            =   495
      TabIndex        =   5
      Top             =   180
      Width           =   2220
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "确认新密码"
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
      Left            =   495
      TabIndex        =   4
      Top             =   585
      Width           =   2220
   End
End
Attribute VB_Name = "ConfirmLogin"
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

Private Sub Cmd_Ok_Click()
    
On Error GoTo Ok_Error
    
    Dim sQuery As String
    Dim md5_pwd As String
    Dim AdoRs As ADODB.Recordset
    
    Set AdoRs = New ADODB.Recordset
    
    If Trim(txt_newpassword.Text) = "" Then
        Call Gp_MsgBoxDisplay("新密码必须输入", "I")
        Exit Sub
    End If
    
    If Trim(txt_confirm.Text) = "" Then
        Call Gp_MsgBoxDisplay("确认新密码必须输入", "I")
        Exit Sub
    End If
    
    If Trim(txt_newpassword.Text) = Trim(txt_confirm.Text) Then
        
        Login.txt_password.Text = ""
        
        'Db Connection Check
        If M_CN1 Is Nothing Then
            If GF_DbConnect = False Then Exit Sub
        End If
        
        sQuery = "SELECT gf_md5_pwd('" + Trim(txt_confirm.Text) + "') FROM dual"
        md5_pwd = Gf_CodeFind(M_CN1, sQuery)
        
        sQuery = "Update ZP_EMPLOYEE set PASSWORD = '" + md5_pwd + "' "
        sQuery = sQuery + " where EMP_ID = '" + sUserID + "' "
        
        'Ado Execute
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
        'AdoRs.Close
        Set AdoRs = Nothing
        Unload Me
        
    Else
        Call Gp_MsgBoxDisplay("新密码和确认新密码输入不一致", "I")
    End If
    
    Set AdoRs = Nothing
    Exit Sub
    
Ok_Error:

    Call Gp_MsgBoxDisplay("密码修改失败...")
    Set AdoRs = Nothing
    
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
