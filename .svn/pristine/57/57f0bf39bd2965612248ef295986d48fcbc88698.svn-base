VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form frmSampStd 
   Caption         =   "取样代码检索 "
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   7935
   StartUpPosition =   3  '窗口缺省
   Begin Threed.SSPanel SSPanel1 
      Height          =   3675
      Left            =   120
      TabIndex        =   2
      Top             =   180
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   6482
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSFrame SSFrame1 
         Height          =   3075
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   5424
         _Version        =   196609
         Caption         =   "取样代码"
         Begin VB.ComboBox cbo_Name1 
            Height          =   300
            Left            =   2220
            TabIndex        =   13
            Top             =   390
            Width           =   5205
         End
         Begin VB.ComboBox cbo_Name2 
            Height          =   300
            Left            =   2220
            TabIndex        =   12
            Top             =   930
            Width           =   5205
         End
         Begin VB.ComboBox cbo_Name3 
            Height          =   300
            Left            =   2220
            TabIndex        =   11
            Top             =   1470
            Width           =   5205
         End
         Begin VB.ComboBox cbo_Name4 
            Height          =   300
            Left            =   2220
            TabIndex        =   10
            Top             =   2010
            Width           =   5175
         End
         Begin VB.ComboBox cbo_Name5 
            Height          =   300
            Left            =   2220
            TabIndex        =   9
            Top             =   2550
            Width           =   5175
         End
         Begin VB.ComboBox cbo_Cd1 
            Height          =   300
            Left            =   7320
            TabIndex        =   8
            Top             =   390
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.ComboBox cbo_Cd2 
            Height          =   300
            Left            =   7320
            TabIndex        =   7
            Top             =   930
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.ComboBox cbo_Cd3 
            Height          =   300
            Left            =   7320
            TabIndex        =   6
            Top             =   1470
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.ComboBox cbo_Cd4 
            Height          =   300
            Left            =   7320
            TabIndex        =   5
            Top             =   2010
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.ComboBox cbo_Cd5 
            Height          =   300
            Left            =   7320
            TabIndex        =   4
            Top             =   2550
            Visible         =   0   'False
            Width           =   555
         End
         Begin InDate.ULabel ULabel3 
            Height          =   300
            Index           =   0
            Left            =   270
            Top             =   390
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            Caption         =   $"frmSampStd.frx":0000
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
         End
         Begin InDate.ULabel ULabel3 
            Height          =   300
            Index           =   1
            Left            =   270
            Top             =   930
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            Caption         =   "长度方向部位 "
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
         End
         Begin InDate.ULabel ULabel3 
            Height          =   300
            Index           =   2
            Left            =   270
            Top             =   1470
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            Caption         =   "宽度方向部位"
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
         End
         Begin InDate.ULabel ULabel3 
            Height          =   300
            Index           =   3
            Left            =   270
            Top             =   2010
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            Caption         =   "厚度方向部位"
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
         End
         Begin InDate.ULabel ULabel3 
            Height          =   300
            Index           =   4
            Left            =   270
            Top             =   2550
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            Caption         =   $"frmSampStd.frx":000E
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
         End
      End
   End
   Begin Threed.SSCommand cmd_Confirm 
      Height          =   315
      Left            =   3630
      TabIndex        =   0
      Top             =   4050
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      _Version        =   196609
      Caption         =   "确定"
      BevelWidth      =   1
   End
   Begin Threed.SSCommand cmd_Cancel 
      Height          =   315
      Left            =   4845
      TabIndex        =   1
      Top             =   4050
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      _Version        =   196609
      Caption         =   "取消"
      BevelWidth      =   1
   End
End
Attribute VB_Name = "frmSampStd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       质量管理
'-- Sub_System Name   质量标准管理
'-- Program Name      取样代码检索
'-- Program ID        AQA0010C
'-- Document No       Q-00-0010(Specification)
'-- Designer          CHU KYO SU
'-- Coder             CHU KYO SU
'-- Date              2003.10.07
'-- Description       取样代码检索
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------

Private Sub Form_Load()
    
    Call Gp_SetSampleCode(frmSampStd)
    
    If Len(sSampSearch) > 1 Then Call subComboSelect
    
End Sub

'试样个数 ComboBox Click
Private Sub cbo_Name1_Click()
    cbo_Cd1.ListIndex = cbo_Name1.ListIndex
End Sub

'长度方向部位  ComboBox Click
Private Sub cbo_Name2_Click()
    cbo_Cd2.ListIndex = cbo_Name2.ListIndex
End Sub

'宽度方向部位 ComboBox Click
Private Sub cbo_Name3_Click()
    cbo_Cd3.ListIndex = cbo_Name3.ListIndex
End Sub

'厚度方向部位 ComboBox Click
Private Sub cbo_Name4_Click()
    cbo_Cd4.ListIndex = cbo_Name4.ListIndex
End Sub

'试样尺寸代码 ComboBox Click
Private Sub cbo_Name5_Click()
    cbo_Cd5.ListIndex = cbo_Name5.ListIndex
End Sub


'试样个数 ComboBox Click
Private Sub cbo_Cd1_Click()
    cbo_Name1.ListIndex = cbo_Cd1.ListIndex
End Sub

'长度方向部位  ComboBox Click
Private Sub cbo_Cd2_Click()
    cbo_Name2.ListIndex = cbo_Cd2.ListIndex
End Sub

'宽度方向部位 ComboBox Click
Private Sub cbo_Cd3_Click()
    cbo_Name3.ListIndex = cbo_Cd3.ListIndex
End Sub

'厚度方向部位 ComboBox Click
Private Sub cbo_Cd4_Click()
    cbo_Name4.ListIndex = cbo_Cd4.ListIndex
End Sub

'试样尺寸代码 ComboBox Click
Private Sub cbo_Cd5_Click()
    cbo_Name5.ListIndex = cbo_Cd5.ListIndex
End Sub

'确定 Button Click
Private Sub cmd_Confirm_Click()

    If fun_Form_Check = False Then Exit Sub
    
    sSampCd = cbo_Cd1.Text + cbo_Cd2.Text + cbo_Cd3.Text + cbo_Cd4.Text + cbo_Cd5.Text
    
    Unload Me

End Sub

'取消 Button Click
Private Sub cmd_Cancel_Click()
    sSampCd = ""
    Unload Me
End Sub


'Form Check
Private Function fun_Form_Check() As Boolean
    
    Dim sMesg As String
    Dim iCnt As Integer
    
    If Trim(cbo_Name1.Text) = "" Then
        iCnt = iCnt + 1
        sMesg = "(试样个数) "
    End If
    
    If Trim(cbo_Name2.Text) = "" Then
        iCnt = iCnt + 1
        sMesg = sMesg + "(长度方向部位 ) "
    End If
    
    If Trim(cbo_Name3.Text) = "" Then
        iCnt = iCnt + 1
        sMesg = sMesg + "(宽度方向部位) "
    End If
    
    If Trim(cbo_Name4.Text) = "" Then
        iCnt = iCnt + 1
        sMesg = sMesg + "(厚度方向部位) "
    End If
    
    If Trim(cbo_Name5.Text) = "" Then
        iCnt = iCnt + 1
        sMesg = sMesg + "(试样尺寸代码) "
    End If
    
    If iCnt > 0 Then GoTo err_form
    
    fun_Form_Check = True
        
    Exit Function
        
err_form:
    sMesg = sMesg + " Must input necessarily"
    Call Gp_MsgBoxDisplay(sMesg)
    
End Function

'ComboBox Init
Private Sub subComboSelect()

    Dim sCode(5) As String
    
    sCode(1) = UCase(Mid(sSampSearch, 1, 2))
    sCode(2) = UCase(Mid(sSampSearch, 3, 1))
    sCode(3) = UCase(Mid(sSampSearch, 4, 1))
    sCode(4) = UCase(Mid(sSampSearch, 5, 1))
    sCode(5) = UCase(Mid(sSampSearch, 6, 2))
        
    If Len(sCode(1)) = 2 Then Call Gp_SetComboBoxListIndex(cbo_Cd1, sCode(1))
        
    If Len(sCode(2)) = 1 Then Call Gp_SetComboBoxListIndex(cbo_Cd2, sCode(2))
        
    If Len(sCode(3)) = 1 Then Call Gp_SetComboBoxListIndex(cbo_Cd3, sCode(3))
    
    If Len(sCode(4)) = 1 Then Call Gp_SetComboBoxListIndex(cbo_Cd4, sCode(4))
    
    If Len(sCode(5)) = 2 Then Call Gp_SetComboBoxListIndex(cbo_Cd5, sCode(5))

End Sub



