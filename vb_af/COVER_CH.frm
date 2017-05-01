VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form COVER_CH 
   BackColor       =   &H00E0E0E0&
   Caption         =   "精炼炉盖更换_COVER_CH"
   ClientHeight    =   2235
   ClientLeft      =   5355
   ClientTop       =   4530
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4605
   Begin VB.TextBox txt_lf_cover3 
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
      Left            =   1830
      TabIndex        =   11
      Top             =   120
      Width           =   855
   End
   Begin Threed.SSCheck CHK_LF_COVER1 
      Height          =   285
      Left            =   360
      TabIndex        =   6
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      _Version        =   196609
      Font3D          =   1
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "#1LF 炉盖更换"
   End
   Begin VB.TextBox txt_lf_cover2 
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
      Left            =   975
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txt_rh_cover 
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
      Left            =   3615
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txt_vd_cover 
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
      Left            =   2730
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txt_lf_cover1 
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
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin Threed.SSCommand cmd_ok 
      Height          =   435
      Left            =   1170
      TabIndex        =   4
      Top             =   1740
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
   Begin Threed.SSCommand cmd_exit 
      Height          =   435
      Left            =   2310
      TabIndex        =   5
      Top             =   1740
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
   Begin Threed.SSCheck CHK_LF_COVER2 
      Height          =   285
      Left            =   360
      TabIndex        =   7
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      _Version        =   196609
      Font3D          =   1
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "#2LF 炉盖更换"
   End
   Begin Threed.SSCheck CHK_VD_COVER 
      Height          =   285
      Left            =   2310
      TabIndex        =   8
      Top             =   780
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      _Version        =   196609
      Font3D          =   1
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " VD 炉盖更换"
   End
   Begin Threed.SSCheck CHK_RH_COVER 
      Height          =   285
      Left            =   2310
      TabIndex        =   9
      Top             =   1290
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      _Version        =   196609
      Font3D          =   1
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "RH 炉盖更换"
   End
   Begin Threed.SSCheck CHK_LF_COVER3 
      Height          =   285
      Left            =   360
      TabIndex        =   10
      Top             =   1320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      _Version        =   196609
      Font3D          =   1
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "#3LF 炉盖更换"
   End
End
Attribute VB_Name = "COVER_CH"
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
'-- Program Name      MPS
'-- Program ID        AFE5012C
'-- Document No
'-- Designer          YANGSHU
'-- Coder             YANGSHU
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
     
   Call Gp_Ms_Collection(txt_lf_cover1, "P", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_lf_cover2, "P", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_lf_cover3, "P", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_vd_cover, "P", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_rh_cover, "P", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         
     Mc1.Add Item:="AFE5012C.P_SREFER", Key:="P-R"
     Mc1.Add Item:="AFE5012C.P_MODIFY", Key:="P-M"
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

Private Sub CHK_LF_COVER1_Click(VALUE As Integer)
   If CHK_LF_COVER1.VALUE = 1 Then
      txt_lf_cover1.Text = "LFC"
   Else
      txt_lf_cover1.Text = ""
   End If
End Sub

Private Sub CHK_LF_COVER2_Click(VALUE As Integer)
   If CHK_LF_COVER2.VALUE = 1 Then
      txt_lf_cover2.Text = "LFC"
   Else
      txt_lf_cover2.Text = ""
   End If
End Sub

Private Sub CHK_LF_COVER3_Click(VALUE As Integer)
   If CHK_LF_COVER3.VALUE = 1 Then
      txt_lf_cover3.Text = "LFC"
   Else
      txt_lf_cover3.Text = ""
   End If
End Sub

Private Sub CHK_VD_COVER_Click(VALUE As Integer)
  If CHK_VD_COVER.VALUE = 1 Then
     txt_vd_cover.Text = "VDC"
   Else
      txt_vd_cover.Text = ""
   End If
End Sub

Private Sub CHK_RH_COVER_Click(VALUE As Integer)
  If CHK_RH_COVER.VALUE = 1 Then
     txt_rh_cover.Text = "RHC"
   Else
      txt_rh_cover.Text = ""
   End If
End Sub

Private Sub Cmd_Ok_Click()
         
    If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then
'       Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
       Call AFE5010C.Form_Ref
        Unload Me
    End If
    
End Sub

Private Sub cmd_exit_Click()
   Call Form_Exit
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

Public Sub Form_Exit()
    Unload Me
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

'    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, AFE5010C.sAuthority)

End Sub
