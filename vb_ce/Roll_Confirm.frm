VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form Roll_Confirm 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "指示确定"
   ClientHeight    =   1920
   ClientLeft      =   6075
   ClientTop       =   4125
   ClientWidth     =   3945
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   3945
   Begin Threed.SSPanel pnl_first 
      Height          =   960
      Left            =   45
      TabIndex        =   3
      Top             =   90
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   1693
      _Version        =   196609
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
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.ComboBox cbo_first 
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
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   450
         Width           =   3390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   135
         Width           =   105
      End
   End
   Begin Threed.SSCommand cmd_OK 
      Height          =   465
      Left            =   630
      TabIndex        =   1
      Top             =   1215
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   820
      _Version        =   196609
      Font3D          =   2
      ForeColor       =   255
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "確認"
   End
   Begin Threed.SSCommand cmd_Cancel 
      Height          =   465
      Left            =   1980
      TabIndex        =   2
      Top             =   1215
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   820
      _Version        =   196609
      Font3D          =   2
      ForeColor       =   16711680
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "取消"
   End
End
Attribute VB_Name = "Roll_Confirm"
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
'-- Program Name      INSTRUCTION CONFIRM
'-- Program ID        Roll_CONFIRM
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
'-- Date              2003.10.8
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
Public P_MODE As String
Public P_PLT As String
Public P_LINE As String
Public P_CurrentCol As Integer

Private Sub Cmd_Cancel_Click()
    Unload Me
End Sub

Private Sub Cmd_Ok_Click()
    
    If cbo_first.Text = "" Then
        Call Gp_MsgBoxDisplay(Label1.Caption & " is must input necessarily")
        Exit Sub
    End If
    
    Call Gp_Process_Exec
    
End Sub

Private Sub Form_Activate()
    
    Call Gp_Combo_Add(Active_Spread)
    
    Select Case P_MODE
        Case "C"    'CAST
            Label1.Caption = "Cast No"
        Case "R"    'Roll
            Label1.Caption = "Roll No"
    End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()

    Call Gp_FormCenter(Me)
    Me.BackColor = &HE0E0E0

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Set Active_Spread = Nothing
    
End Sub

Public Sub Gp_Combo_Add(sPname As Variant)

    Dim iRow As Integer
    Dim sTemp As String
    
    With sPname
    
        .Col = P_CurrentCol
        cbo_first.AddItem ""
                
        For iRow = 1 To .MaxRows
            .Row = iRow
            
            If sTemp <> .Text Then
                cbo_first.AddItem .Text
                sTemp = .Text
            End If
                    
            If iRow = 1 Then sTemp = .Text
            
        Next iRow
        
    End With
    
End Sub


Public Sub Gp_Process_Exec()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    
    Dim adoCmd As ADODB.Command
    
    'Exit Sub '-------------------------
    
    Screen.MousePointer = vbHourglass
    
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    If P_MODE = "C" Then 'CAST   SLAB_EDT_FL = '1'  HCR
        'sQuery = "{call AEE1000P ('" + P_PLT + "','" + P_LINE + "','','" + Trim(cbo_first.Text) + "','','" + sUserID + "',?)}"
    Else                 'ROLL   SLAB_EDT_FL = '2'  CCR
        sQuery = "{call CED1100P ('" + P_PLT + "','" + P_LINE + "','','','" + Trim(cbo_first.Text) + "','" + sUserID + "',?)}"
    End If
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
    
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    
    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    If adoCmd("arg_e_msg") <> "" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Screen.MousePointer = vbDefault
        Call Gp_MsgBoxDisplay(sErrMessg)
        Set adoCmd = Nothing
        Exit Sub
    Else
        Call Gp_MsgBoxDisplay("指示确定完了..!!", "I")
        If P_MODE = "C" Then
            'AEC1070C.Complete = True
        Else
            CEC2900C.Complete = True
        End If
        Set adoCmd = Nothing
        Screen.MousePointer = vbDefault
        Unload Me
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_ERROR : " & Error)
    
End Sub

