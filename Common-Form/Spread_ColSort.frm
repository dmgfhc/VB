VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form Spread_ColSort 
   Caption         =   "表格排序"
   ClientHeight    =   3915
   ClientLeft      =   4680
   ClientTop       =   3750
   ClientWidth     =   5355
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   5355
   Begin Threed.SSPanel pnl_first 
      Height          =   960
      Left            =   45
      TabIndex        =   2
      Top             =   90
      Width           =   5235
      _ExtentX        =   9234
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
      Begin Threed.SSOption opt_first_a 
         Height          =   285
         Left            =   3780
         TabIndex        =   5
         Top             =   180
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   2
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
         Caption         =   "升序"
         Value           =   -1
      End
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
         TabIndex        =   4
         Top             =   450
         Width           =   3390
      End
      Begin Threed.SSOption opt_first_d 
         Height          =   285
         Left            =   3780
         TabIndex        =   6
         Top             =   495
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   2
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
         Caption         =   "降序"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "第一排序列"
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
         TabIndex        =   3
         Top             =   135
         Width           =   975
      End
   End
   Begin Threed.SSCommand cmd_OK 
      Height          =   465
      Left            =   1395
      TabIndex        =   0
      Top             =   3285
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
      Caption         =   "确认"
   End
   Begin Threed.SSCommand cmd_Cancel 
      Height          =   465
      Left            =   2745
      TabIndex        =   1
      Top             =   3285
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
   Begin Threed.SSPanel SSPanel1 
      Height          =   960
      Left            =   45
      TabIndex        =   7
      Top             =   1125
      Width           =   5235
      _ExtentX        =   9234
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
      Begin VB.ComboBox cbo_second 
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
         TabIndex        =   9
         Top             =   450
         Width           =   3390
      End
      Begin Threed.SSOption opt_second_a 
         Height          =   285
         Left            =   3780
         TabIndex        =   8
         Top             =   180
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   2
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
         Caption         =   "升序"
         Value           =   -1
      End
      Begin Threed.SSOption opt_second_d 
         Height          =   285
         Left            =   3780
         TabIndex        =   10
         Top             =   495
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   2
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
         Caption         =   "降序"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "第二排序列"
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
         TabIndex        =   11
         Top             =   135
         Width           =   975
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   960
      Left            =   45
      TabIndex        =   12
      Top             =   2160
      Width           =   5235
      _ExtentX        =   9234
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
      Begin VB.ComboBox cbo_third 
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
         TabIndex        =   13
         Top             =   450
         Width           =   3390
      End
      Begin Threed.SSOption opt_third_a 
         Height          =   285
         Left            =   3780
         TabIndex        =   14
         Top             =   180
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   2
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
         Caption         =   "升序"
         Value           =   -1
      End
      Begin Threed.SSOption opt_third_d 
         Height          =   285
         Left            =   3780
         TabIndex        =   15
         Top             =   495
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   2
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
         Caption         =   "降序"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "第三排序列"
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
         TabIndex        =   16
         Top             =   135
         Width           =   975
      End
   End
End
Attribute VB_Name = "Spread_ColSort"
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
'-- Program Name      Multi Column Sort
'-- Program ID        Spread_ColSort
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
'-- Date              2003.5.19
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------

Private Sub cbo_first_Click()
    
    If cbo_first.Text = "" Then
        
        cbo_Second.ListIndex = 0
        cbo_third.ListIndex = 0
        
        cbo_Second.Enabled = False
        cbo_third.Enabled = False
        
        opt_second_a.Enabled = False
        opt_second_d.Enabled = False
    
        opt_third_a.Enabled = False
        opt_third_d.Enabled = False
        
        Exit Sub
        
    Else
        
        cbo_Second.Enabled = True
        
        opt_second_a.Enabled = True
        opt_second_d.Enabled = True
        
    End If
    
    If cbo_first.Text = cbo_Second.Text Or cbo_first.Text = cbo_third.Text Then
        
        Call Gp_MsgBoxDisplay("该列已被选择", "I")
        
        cbo_first.ListIndex = 0
    
    End If
    
End Sub

Private Sub cbo_second_click()

    If cbo_Second.Text = "" Then
        
        cbo_third.ListIndex = 0
        cbo_third.Enabled = False
        
        opt_third_a.Enabled = False
        opt_third_d.Enabled = False
        
        Exit Sub
    Else
        
        cbo_third.Enabled = True
        
        opt_third_a.Enabled = True
        opt_third_d.Enabled = True
        
    End If
    
    If cbo_Second.Text = cbo_first.Text Or cbo_Second.Text = cbo_third.Text Then
        
        Call Gp_MsgBoxDisplay("该列已被选择", "I")
        
        cbo_Second.ListIndex = 0
    
    End If

End Sub

Private Sub cbo_third_click()

    If cbo_third.Text = "" Then Exit Sub
    
    If cbo_third.Text = cbo_first.Text Or cbo_third.Text = cbo_Second.Text Then
        
        Call Gp_MsgBoxDisplay("该列已被选择", "I")
        
        cbo_third.ListIndex = 0
    
    End If

End Sub

Private Sub Cmd_Cancel_Click()
    Unload Me
End Sub

Private Sub Cmd_Ok_Click()

    With Active_Spread
    
        If cbo_first.Text <> "" Then
        
            .SortKey(1) = CInt(Right(cbo_first.Text, 2))        'col
            
            If opt_first_a.Value = True Then
                .SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
            Else
                .SortKeyOrder(1) = SS_SORT_ORDER_DESCENDING
            End If
            
        End If
        
        If cbo_Second.Text <> "" Then
        
            .SortKey(2) = CInt(Right(cbo_Second.Text, 2))        'col
            
            If opt_second_a.Value = True Then
                .SortKeyOrder(2) = SS_SORT_ORDER_ASCENDING
            Else
                .SortKeyOrder(2) = SS_SORT_ORDER_DESCENDING
            End If
            
        End If
        
        If cbo_third.Text <> "" Then
        
            .SortKey(3) = CInt(Right(cbo_third.Text, 2))        'col
            
            If opt_third_a.Value = True Then
                .SortKeyOrder(3) = SS_SORT_ORDER_ASCENDING
            Else
                .SortKeyOrder(3) = SS_SORT_ORDER_DESCENDING
            End If
            
        End If
        
        .Sort 0, 1, .MaxCols, .MaxRows, SortByRow

    End With

    Unload Me
    
End Sub

Private Sub Form_Load()

    Call Gp_Sp_ColSort(Active_Spread, Me)
    Call Gp_FormCenter(Me)
    
    cbo_Second.Enabled = False
    cbo_third.Enabled = False
    
    opt_second_a.Enabled = False
    opt_second_d.Enabled = False
    
    opt_third_a.Enabled = False
    opt_third_d.Enabled = False
    
    Me.BackColor = &HE0E0E0
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set Active_Spread = Nothing
    
End Sub

Private Sub opt_first_a_Click(Value As Integer)
    
    If opt_first_a.Value = True Then
        opt_first_a.ForeColor = &HFF&
        opt_first_d.ForeColor = &H808080
    Else
        opt_first_d.ForeColor = &HFF&
        opt_first_a.ForeColor = &H808080
    End If
    
End Sub

Private Sub opt_first_d_Click(Value As Integer)

    If opt_first_d.Value = True Then
        opt_first_d.ForeColor = &HFF&
        opt_first_a.ForeColor = &H808080
    Else
        opt_first_a.ForeColor = &HFF&
        opt_first_d.ForeColor = &H808080
    End If

End Sub

Private Sub opt_second_a_Click(Value As Integer)

    If opt_second_a.Value = True Then
        opt_second_a.ForeColor = &HFF&
        opt_second_d.ForeColor = &H808080
    Else
        opt_second_d.ForeColor = &HFF&
        opt_second_a.ForeColor = &H808080
    End If

End Sub

Private Sub opt_second_d_Click(Value As Integer)

    If opt_second_d.Value = True Then
        opt_second_d.ForeColor = &HFF&
        opt_second_a.ForeColor = &H808080
    Else
        opt_second_a.ForeColor = &HFF&
        opt_second_d.ForeColor = &H808080
    End If

End Sub

Private Sub opt_third_a_Click(Value As Integer)

    If opt_third_a.Value = True Then
        opt_third_a.ForeColor = &HFF&
        opt_third_d.ForeColor = &H808080
    Else
        opt_third_d.ForeColor = &HFF&
        opt_third_a.ForeColor = &H808080
    End If

End Sub

Private Sub opt_third_d_Click(Value As Integer)

    If opt_third_d.Value = True Then
        opt_third_d.ForeColor = &HFF&
        opt_third_a.ForeColor = &H808080
    Else
        opt_third_a.ForeColor = &HFF&
        opt_third_d.ForeColor = &H808080
    End If

End Sub
