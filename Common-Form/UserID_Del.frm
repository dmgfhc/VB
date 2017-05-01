VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form UserID_Del 
   BackColor       =   &H00FFFAE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UserID_Del"
   ClientHeight    =   1590
   ClientLeft      =   5475
   ClientTop       =   8805
   ClientWidth     =   5880
   Icon            =   "UserID_Del.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   5880
   Begin Threed.SSFrame SSFrame1 
      Height          =   1365
      Left            =   180
      TabIndex        =   3
      Top             =   90
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   2408
      _Version        =   196609
      Font3D          =   1
      BackColor       =   16775907
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "删除用户代码"
      ShadowStyle     =   1
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
         Left            =   2160
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   315
         Width           =   2310
      End
      Begin Threed.SSCommand cmd_Cancel 
         Height          =   420
         Left            =   2790
         TabIndex        =   2
         Tag             =   "取消"
         Top             =   810
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
         Left            =   1440
         TabIndex        =   1
         Tag             =   "確認"
         Top             =   810
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
         Left            =   1035
         TabIndex        =   4
         Top             =   360
         Width           =   1050
      End
   End
End
Attribute VB_Name = "UserID_Del"
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
'-- Program Name      UserID_Del
'-- Program ID        UserID_Del
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
'-- Date              2003.12.5
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

    Dim Reg_id As Variant
    Dim intSettings As Integer
    
    'Registry ID Delete
    DeleteSetting "NISCO", "LOGIN", cbo_emp_id.Text
    
    Call Gp_MsgBoxDisplay("用户代码删除成功..", "I")
    
    'Registry ID Setting
    Reg_id = GetAllSettings("NISCO", "LOGIN")
    
    cbo_emp_id.Clear
    
    For intSettings = LBound(Reg_id, 1) To UBound(Reg_id, 1)
        If Reg_id(intSettings, 0) <> "SAMPLE" Then
            cbo_emp_id.AddItem Reg_id(intSettings, 0)
        End If
    Next intSettings
    
    Exit Sub
    
Ok_Error:

    Call Gp_MsgBoxDisplay("用户代码删除失败...")
    
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

    Dim Reg_id As Variant
    Dim intSettings As Integer
    
    Me.KeyPreview = True
    'Me.BackColor = &HE0E0E0
    
    Call Gp_FormCenter(Me)
    
    'Registry ID Setting
    Reg_id = GetAllSettings("NISCO", "LOGIN")
    
    cbo_emp_id.Clear
    
    For intSettings = LBound(Reg_id, 1) To UBound(Reg_id, 1)
        If Reg_id(intSettings, 0) <> "SAMPLE" Then
            cbo_emp_id.AddItem Reg_id(intSettings, 0)
        End If
    Next intSettings
    
End Sub
