VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form Process_Change 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Process_Change"
   ClientHeight    =   2670
   ClientLeft      =   3210
   ClientTop       =   3420
   ClientWidth     =   6105
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   6105
   Begin Threed.SSPanel pnl_first 
      Height          =   1365
      Left            =   45
      TabIndex        =   2
      Top             =   585
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   2408
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
      Begin VB.ComboBox cbo_target 
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
         Left            =   855
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   855
         Width           =   2265
      End
      Begin VB.ComboBox cbo_to 
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
         Left            =   3195
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   450
         Width           =   2265
      End
      Begin VB.ComboBox cbo_from 
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
         Left            =   855
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   450
         Width           =   2265
      End
      Begin Threed.SSOption opt_top 
         Height          =   285
         Left            =   3555
         TabIndex        =   13
         Top             =   900
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   255
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
         Caption         =   "前"
         Value           =   -1
      End
      Begin Threed.SSOption opt_bottom 
         Height          =   285
         Left            =   4320
         TabIndex        =   14
         Top             =   900
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   8421504
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
         Caption         =   "后"
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "目标"
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
         Left            =   300
         TabIndex        =   12
         Top             =   900
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "对象"
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
         Left            =   300
         TabIndex        =   9
         Top             =   495
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "No"
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
         Left            =   300
         TabIndex        =   3
         Top             =   135
         Width           =   210
      End
   End
   Begin Threed.SSCommand cmd_OK 
      Height          =   465
      Left            =   1680
      TabIndex        =   0
      Top             =   2070
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   820
      _Version        =   196609
      Font3D          =   1
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
      Left            =   3210
      TabIndex        =   1
      Top             =   2070
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   820
      _Version        =   196609
      Font3D          =   1
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
   Begin Threed.SSPanel SSPanel1 
      Height          =   510
      Left            =   45
      TabIndex        =   5
      Top             =   45
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   900
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
      Begin Threed.SSOption opt_move 
         Height          =   285
         Left            =   945
         TabIndex        =   6
         Top             =   135
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   255
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
         Caption         =   "移动"
         Value           =   -1
      End
      Begin Threed.SSOption opt_split 
         Height          =   285
         Left            =   3150
         TabIndex        =   7
         Top             =   135
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   8421504
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
         Caption         =   "分开"
      End
      Begin Threed.SSOption opt_unification 
         Height          =   285
         Left            =   2055
         TabIndex        =   8
         Top             =   135
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   8421504
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
         Caption         =   "统合"
      End
      Begin Threed.SSOption opt_delete 
         Height          =   285
         Left            =   4200
         TabIndex        =   15
         Top             =   135
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   8421504
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
         Caption         =   "删除"
      End
   End
End
Attribute VB_Name = "Process_Change"
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
'-- Program Name      Porcess Change
'-- Program ID        Process_Change
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
'-- Date              2003.6.27
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------

Public P_PLT As String              'PLT
Public P_LINE As Integer            'LINE = '1'
Public P_STATUS As String           'DAILY = 'D', INSTRUCTION = 'I'
Public P_MODE As String             'MOVE = 'M',  SPLIT = 'S', UNIFICATION = 'U', DELETE = 'D'
Public P_UNIT As String             'PLATE = 'P', SLAB = 'S',  CHARGE = 'H', CAST = 'C', ROLL = 'R'
Public P_POSITION As String         'TOP = 'T',   BOTTOM = 'B'

Public P_Tcurrent As Integer        'Target Column
Public P_CurrentCol As Integer      'PLATE, SLAB, CHARGE, CAST, ROLL Column

Private Sub cbo_from_Click()

    Select Case P_UNIT
        
        Case "R"   'Roll
        
            If P_MODE <> "M" Then
                cbo_to.Text = cbo_from.Text
            End If
        
        Case "S"   'SLAB
        
            If P_MODE = "D" Then
                cbo_to.Text = cbo_from.Text
            End If
            
        Case "C"   'Cast
        
            If P_MODE <> "M" Then
                cbo_to.ListIndex = cbo_from.ListIndex
            End If
    
    End Select
        
End Sub

Private Sub cbo_to_Click()

    Select Case P_UNIT
        
        Case "C"   'Cast
        
            If cbo_to.ListIndex = 0 Then Exit Sub
            If P_MODE = "M" Then
                If Trim(cbo_from.Text) > Trim(cbo_to.Text) Then
                    Call Gp_MsgBoxDisplay("Can not be small than To Cast No a from Cast No")
                    cbo_to.ListIndex = 0
                End If
            End If
            
    End Select
    
End Sub

Private Sub Cmd_Cancel_Click()
    Unload Me
End Sub

Private Sub Cmd_Ok_Click()
    
    If cbo_from.Text = "" Or cbo_to.Text = "" Or cbo_target.Text = "" Then
        If P_MODE = "D" Then
            If cbo_from.Text = "" Or cbo_to.Text = "" Then
                Call Gp_MsgBoxDisplay("Must input Value of From, To item")
                Exit Sub
            End If
        Else
            Call Gp_MsgBoxDisplay("Must input From, To, Value of Target item")
            Exit Sub
        End If
    End If
    
    If cbo_from.Text <> "" And cbo_to.Text <> "" And cbo_target.Text <> "" Then
        If Len(cbo_from.Text) = Len(cbo_to.Text) And Len(cbo_to.Text) = Len(cbo_target.Text) Then
            If cbo_from.Text <= cbo_target.Text And cbo_target.Text <= cbo_from.Text Then
                Call Gp_MsgBoxDisplay("Value of Target item is between from and to..")
                Exit Sub
            End If
        End If
    End If
    
    Call Gp_Process_Exec
    
End Sub

Private Sub Form_Activate()
    
    P_MODE = "M"
    P_POSITION = "T"
    
    Call Gp_Combo_Add(Active_Spread)
    
    Select Case P_UNIT
    
        Case "H"    'Charge
            opt_unification.Enabled = False
            opt_split.Enabled = False
            Label1.Caption = "Charge No"
        Case "S"    'Slab
            opt_unification.Enabled = False
            opt_split.Enabled = False
            Label1.Caption = "Slab No"
        Case "P"    'Plate
            opt_unification.Enabled = False
            opt_split.Enabled = False
            Label1.Caption = "Plate No"
        Case "C"    'Cast
            opt_delete.Enabled = False
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

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Set Active_Spread = Nothing

End Sub

Public Sub Gp_Combo_Add(sPname As Variant)

    Dim iRow As Integer
    Dim sTemp As String
    
    With sPname
    
        .Col = P_CurrentCol
        
        cbo_from.AddItem ""
        cbo_to.AddItem ""
        cbo_target.AddItem ""
                
        For iRow = 1 To .MaxRows
            .Row = iRow
            
            If sTemp <> .Text Then
            
                cbo_from.AddItem .Text
                cbo_to.AddItem .Text
                cbo_target.AddItem .Text
                sTemp = .Text
            
            End If
                    
            If iRow = 1 Then sTemp = .Text
            
        Next iRow
        
        Select Case P_UNIT
            
            Case "R"   'Roll   ------ Target Combo Re-setting
            
                cbo_target.Clear
                
                If P_MODE = "M" Then
                    .Col = 1
                Else
                    .Col = P_Tcurrent
                End If
                
                For iRow = 1 To .MaxRows
                    .Row = iRow
                    
                    If sTemp <> .Text Then
                        cbo_target.AddItem .Text
                        sTemp = .Text
                    End If
                            
                    If iRow = 1 Then sTemp = .Text
                    
                Next iRow
                
        End Select
    
    End With
    
End Sub

Private Sub opt_bottom_Click(Value As Integer)

    If opt_bottom.Value = True Then
        opt_bottom.ForeColor = &HFF&
        opt_top.ForeColor = &H808080
        P_POSITION = "B"
    Else
        opt_bottom.ForeColor = &H808080
        P_POSITION = "T"
    End If

End Sub

Private Sub opt_delete_Click(Value As Integer)

    Dim iRow As Integer
    Dim sTemp As String

    If opt_delete.Value = True Then
        opt_delete.ForeColor = &HFF&
        opt_unification.ForeColor = &H808080
        opt_split.ForeColor = &H808080
        opt_move.ForeColor = &H808080
        
        P_MODE = "D"
        
        With Active_Spread
        
'            If P_UNIT = "C" Then
'
'                cbo_target.Clear
'
'                .Col = P_CurrentCol
'
'                For iRow = 1 To .MaxRows
'                    .Row = iRow
'
'                    If sTemp <> .Text Then
'                        cbo_target.AddItem .Text
'                        sTemp = .Text
'                    End If
'
'                    If iRow = 1 Then sTemp = .Text
'
'                Next iRow
'
'                cbo_to.ListIndex = cbo_from.ListIndex
'
'            End If
                
        End With
        
        cbo_to.Enabled = True
        
        cbo_target.Enabled = False
        opt_top.Enabled = False
        opt_bottom.Enabled = False
        
        P_POSITION = "T"
        
        If P_UNIT = "S" Then
            cbo_to.ListIndex = cbo_from.ListIndex
            cbo_target.Clear
            cbo_target.Enabled = False
        End If
        
        If P_UNIT = "R" Then
            cbo_to.ListIndex = cbo_from.ListIndex
            cbo_to.Enabled = False
            cbo_target.Clear
            cbo_target.Enabled = False
        End If
        
    Else
        opt_move.ForeColor = &H808080
    End If
    
End Sub

Private Sub opt_move_Click(Value As Integer)

    Dim iRow As Integer
    Dim sTemp As String

    If opt_move.Value = True Then
        opt_move.ForeColor = &HFF&
        opt_unification.ForeColor = &H808080
        opt_split.ForeColor = &H808080
        opt_delete.ForeColor = &H808080
        
        P_MODE = "M"
        
        With Active_Spread
        
            If P_UNIT = "C" Or P_UNIT = "R" Or P_UNIT = "S" Or P_UNIT = "H" Then
                
                cbo_target.Enabled = True
                cbo_target.Clear
                
                If P_UNIT = "R" Then
                    .Col = 1
                Else
                    .Col = P_CurrentCol
                End If
                
                For iRow = 1 To .MaxRows
                    .Row = iRow
                    
                    If sTemp <> .Text Then
                        cbo_target.AddItem .Text
                        sTemp = .Text
                    End If
                            
                    If iRow = 1 Then sTemp = .Text
                    
                Next iRow
            End If
                
        End With
        
        cbo_to.Enabled = True
        
        opt_bottom.Enabled = True
        opt_top.Enabled = True
        opt_top.Value = True
        P_POSITION = "T"
                
    Else
        opt_move.ForeColor = &H808080
    End If

End Sub

Private Sub opt_split_Click(Value As Integer)

    Dim iRow As Integer
    Dim sTemp As String

    If opt_split.Value = True Then
        opt_split.ForeColor = &HFF&
        opt_move.ForeColor = &H808080
        opt_unification.ForeColor = &H808080
        opt_delete.ForeColor = &H808080
        
        P_MODE = "S"
        
        With Active_Spread
        
            If P_UNIT = "C" Or P_UNIT = "R" Then
            
                cbo_target.Clear
                .Col = P_Tcurrent
                
                For iRow = 1 To .MaxRows
                    .Row = iRow
                    
                    If sTemp <> .Text Then
                        cbo_target.AddItem .Text
                        sTemp = .Text
                    End If
                            
                    If iRow = 1 Then sTemp = .Text
                    
                Next iRow
                
                cbo_to.ListIndex = cbo_from.ListIndex
                
                cbo_target.Enabled = True
                opt_bottom.Enabled = True
                opt_top.Enabled = True
                opt_top.Value = True
                P_POSITION = "T"
            
            End If
            
        End With
        
        cbo_to.Enabled = False
                
    Else
        opt_split.ForeColor = &H808080
    End If

End Sub

Private Sub opt_top_Click(Value As Integer)

    If opt_top.Value = True Then
        opt_top.ForeColor = &HFF&
        opt_bottom.ForeColor = &H808080
        P_POSITION = "T"
    Else
        opt_top.ForeColor = &H808080
        P_POSITION = "B"
    End If

End Sub

Private Sub opt_unification_Click(Value As Integer)

    Dim iRow As Integer
    Dim sTemp As String

    If opt_unification.Value = True Then
        opt_unification.ForeColor = &HFF&
        opt_move.ForeColor = &H808080
        opt_split.ForeColor = &H808080
        opt_delete.ForeColor = &H808080
        
        P_MODE = "U"
        
        With Active_Spread
        
            If P_UNIT = "C" Or P_UNIT = "R" Then
            
                cbo_target.Enabled = True
                cbo_target.Clear
                
                If P_UNIT = "R" Then
                    .Col = 1
                Else
                    .Col = P_Tcurrent
                End If
                
                For iRow = 1 To .MaxRows
                    .Row = iRow
                    
                    If sTemp <> .Text Then
                        cbo_target.AddItem .Text
                        sTemp = .Text
                    End If
                            
                    If iRow = 1 Then sTemp = .Text
                    
                Next iRow
            
                opt_bottom.Enabled = True
                opt_top.Enabled = True
                opt_top.Value = True
                P_POSITION = "T"
            
            End If
            
            cbo_target.Enabled = True
            cbo_to.ListIndex = cbo_from.ListIndex
        
        End With
        
        cbo_to.Enabled = False
        
    Else
        opt_unification.ForeColor = &H808080
    End If

End Sub

Public Sub Gp_Process_Exec()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    
    Dim adoCmd As ADODB.Command
    
    Screen.MousePointer = vbHourglass
    
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    sQuery = "{call CEZ5000P ('" + P_PLT + "','" + Trim(Str(P_LINE)) + "','" + P_STATUS + "','" + P_MODE + "','" + P_UNIT + "','" + Trim(cbo_from.Text) + "','"
    sQuery = sQuery + Trim(cbo_to.Text) + "','" + Trim(cbo_target.Text) + "','" + P_POSITION + "','" + sUserID + "',?)}"
    
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
        If P_UNIT = "H" Or P_UNIT = "C" Then
            'AEC1070C.Complete = True
        Else
            CEC2900C.Complete = True
        End If
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
