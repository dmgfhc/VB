VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form Mltcd_Change 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "修改工艺路径(精炼)"
   ClientHeight    =   3015
   ClientLeft      =   9960
   ClientTop       =   2265
   ClientWidth     =   5025
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5025
   Begin Threed.SSPanel pnl_first 
      Height          =   2115
      Left            =   45
      TabIndex        =   2
      Top             =   90
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   3731
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
      Begin VB.TextBox txt_MLT_PROC_CD3 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3120
         TabIndex        =   10
         Top             =   1245
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox cob_MLT_PROC_CD_3 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "Mltcd_Change.frx":0000
         Left            =   3450
         List            =   "Mltcd_Change.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   570
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.TextBox txt_MLT_PROC_CD_ORG 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2100
         TabIndex        =   8
         Top             =   1620
         Width           =   1545
      End
      Begin VB.TextBox txt_MLT_PROC_CD2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2625
         TabIndex        =   7
         Top             =   1245
         Width           =   495
      End
      Begin VB.ComboBox cob_MLT_PROC_CD_2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "Mltcd_Change.frx":001D
         Left            =   2415
         List            =   "Mltcd_Change.frx":002A
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   570
         Width           =   1005
      End
      Begin VB.ComboBox cob_MLT_PROC_CD_1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "Mltcd_Change.frx":003A
         Left            =   1410
         List            =   "Mltcd_Change.frx":0050
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   570
         Width           =   1005
      End
      Begin VB.TextBox txt_MLT_PROC_CD 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2100
         TabIndex        =   4
         Top             =   1245
         Width           =   495
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   0
         Left            =   330
         Top             =   1245
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   556
         Caption         =   "选择工艺路径代码"
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
         ForeColor       =   255
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   16
         Left            =   345
         Top             =   555
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         Caption         =   "工艺路径"
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
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   1
         Left            =   1395
         Top             =   195
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         Caption         =   "工序 1"
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
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   3
         Left            =   2415
         Top             =   195
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         Caption         =   "工序 2"
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
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   4
         Left            =   330
         Top             =   1620
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   556
         Caption         =   "原工艺路径代码"
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
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   2
         Left            =   3450
         Top             =   195
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         Caption         =   "工序 3"
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
         TabIndex        =   3
         Top             =   45
         Width           =   105
      End
   End
   Begin Threed.SSCommand cmd_OK 
      Height          =   465
      Left            =   1170
      TabIndex        =   0
      Top             =   2460
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
      Left            =   2595
      TabIndex        =   1
      Top             =   2460
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
End
Attribute VB_Name = "Mltcd_Change"
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
'-- Program Name      MLTCD SET AND CHANGE
'-- Program ID        MLTCD_CHANGE
'-- Document No       Q-00-0010(Specification)
'-- Designer          HJD
'-- Coder             HJD
'-- Date              2006.7.21
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------

Public P_PLT As String
Public P_LINE As String
Public P_CurrentCol As Integer

Private Sub Cmd_Cancel_Click()

    Dim iRow As Integer
    
    With Active_Spread
        For iRow = 0 To .MaxRows
            .Row = iRow: .Col = 0
            If .Text = "Selected" Then
                .Text = ""
                Call Gp_Sp_BlockColor(Active_Spread, 1, Active_Spread.MaxCols, iRow, iRow)
            End If
        Next
    End With
    
    With AEC1070C
        .strCCM_CD1_Pre = ""
        .strCCM_CD1 = ""
        .lngCurRow = 0
        .lngPreRow = 0
        .intCount = 0
    End With
    
    Unload Me

End Sub

Private Sub Cmd_Ok_Click()

    Dim strPrc_CD As String
    Dim strOrg_CD As String
    Dim iRow As Integer
    
    strPrc_CD = Trim(txt_MLT_PROC_CD.Text) + Trim(txt_MLT_PROC_CD2.Text)
    
    Screen.MousePointer = vbHourglass
    
    If Len(Trim(strPrc_CD)) > 0 Then
    
        With Active_Spread
            .Col = 0
            
            For iRow = 0 To .MaxRows
                .Row = iRow
                If .Text = "Selected" Then
                    .Col = 9
                   If Mid(Trim(txt_MLT_PROC_CD.Text), 1, 2) = "BD" Then
                        .Text = Replace(.Text, Mid(.Text, InStr(1, .Text, "BD"), 3), Trim(txt_MLT_PROC_CD.Text))
                   ElseIf Mid(Trim(txt_MLT_PROC_CD.Text), 1, 2) = "BH" Then
                        .Text = Replace(.Text, Mid(.Text, InStr(1, .Text, "BE"), 3), Trim(txt_MLT_PROC_CD.Text))
                   ElseIf Mid(Trim(txt_MLT_PROC_CD.Text), 1, 2) = "BE" Then
                        .Text = Replace(.Text, Mid(.Text, InStr(1, .Text, "BH"), 3), Trim(txt_MLT_PROC_CD.Text))
                   End If
                   
                   If Mid(Trim(txt_MLT_PROC_CD2.Text), 1, 2) = "BD" Then
                        .Text = Replace(.Text, Mid(.Text, InStr(1, .Text, "BD"), 3), Trim(txt_MLT_PROC_CD2.Text))
                   ElseIf Mid(Trim(txt_MLT_PROC_CD2.Text), 1, 2) = "BH" Then
                        .Text = Replace(.Text, Mid(.Text, InStr(1, .Text, "BE"), 3), Trim(txt_MLT_PROC_CD2.Text))
                   ElseIf Mid(Trim(txt_MLT_PROC_CD2.Text), 1, 2) = "BE" Then
                        .Text = Replace(.Text, Mid(.Text, InStr(1, .Text, "BH"), 3), Trim(txt_MLT_PROC_CD2.Text))
                   End If

'                    strOrg_CD = Mid(.Text, InStr(1, .Text, "BC"), 3) + Mid(.Text, InStr(1, "BG", 3), 3)
'                    strOrg_CD = strOrg_CD + IIf(Len(Trim(txt_MLT_PROC_CD.Text)) > 0, Trim(txt_MLT_PROC_CD.Text), Mid(.Text, Len(strOrg_CD), 3))
'                    strOrg_CD = strOrg_CD + IIf(Len(Trim(txt_MLT_PROC_CD2.Text)) > 0, Trim(txt_MLT_PROC_CD2.Text), Mid(.Text, Len(strOrg_CD), 3)) + Mid(.Text, InStr(1, .Text, "BF"), 3)
                    Call Proc_CD_Change(Active_Spread, iRow)
                    Call Gp_Sp_BlockColor(Active_Spread, 1, Active_Spread.MaxCols, iRow, iRow)
                    .Col = 0: .Text = ""
                End If
            Next
            
        End With
        
    End If
    
    Call Gp_MsgBoxDisplay("工艺路径修改结束!", "I")
    Call AEC1070C.Form_Ref
    AEC1070C.Complete = True
    
    Screen.MousePointer = vbDefault
    Unload Me
        
End Sub

Private Sub cob_MLT_PROC_CD_1_Change()

    Dim CD As String
    
    If cob_MLT_PROC_CD_1.Text = "CAS" Then
        CD = "BG" + Mid(txt_MLT_PROC_CD_ORG.Text, 3, 1)
    ElseIf cob_MLT_PROC_CD_1.Text = "1# LF" Then
        CD = "BD1"
    ElseIf cob_MLT_PROC_CD_1.Text = "2# LF" Then
        CD = "BD2"
    ElseIf cob_MLT_PROC_CD_1.Text = "3# LF" Then
        CD = "BD3"
    ElseIf cob_MLT_PROC_CD_1.Text = "VD" Then
        CD = "BE1"
    ElseIf cob_MLT_PROC_CD_1.Text = "RH" Then
        CD = "BH2"
    Else
        CD = "   "
    End If
    
    txt_MLT_PROC_CD.Text = Trim(CD)
    
End Sub

'Private Sub cob_MLT_PROC_CD_1_Click()
'    Dim CD As String
'    With cob_MLT_PROC_CD_1
'    Select Case .ListIndex
'        Case 0
'            cob_MLT_PROC_CD_2.Clear
'            Call cob_MLT_PROC_CD_2.AddItem(" ", 0)
'            cob_MLT_PROC_CD_2.ListIndex = 0
'            CD = "   "
'        Case 1
'            cob_MLT_PROC_CD_2.Clear
'            Call cob_MLT_PROC_CD_2.AddItem(" ", 0)
'            Call cob_MLT_PROC_CD_2.AddItem("1# LF", 1)
'            Call cob_MLT_PROC_CD_2.AddItem("2# LF", 2)
'            Call cob_MLT_PROC_CD_2.AddItem("VD", 3)
'            Call cob_MLT_PROC_CD_2.AddItem("RH", 4)
'            cob_MLT_PROC_CD_2.ListIndex = 0
'            CD = "BG" + Mid(txt_MLT_PROC_CD_ORG.Text, 3, 1)
'        Case 2
'            cob_MLT_PROC_CD_2.Clear
'            Call cob_MLT_PROC_CD_2.AddItem(" ", 0)
'            Call cob_MLT_PROC_CD_2.AddItem("VD", 1)
'            Call cob_MLT_PROC_CD_2.AddItem("RH", 2)
'            cob_MLT_PROC_CD_2.ListIndex = 0
'            CD = "BD1"
'        Case 3
'            cob_MLT_PROC_CD_2.Clear
'            Call cob_MLT_PROC_CD_2.AddItem(" ", 0)
'            Call cob_MLT_PROC_CD_2.AddItem("VD", 1)
'            Call cob_MLT_PROC_CD_2.AddItem("RH", 2)
'            cob_MLT_PROC_CD_2.ListIndex = 0
'            CD = "BD2"
'        Case 4
'            cob_MLT_PROC_CD_2.Clear
'            Call cob_MLT_PROC_CD_2.AddItem(" ", 0)
'            Call cob_MLT_PROC_CD_2.AddItem("RH", 1)
'            cob_MLT_PROC_CD_2.ListIndex = 0
'            CD = "BE1"
'        Case 5
'            cob_MLT_PROC_CD_2.Clear
'            Call cob_MLT_PROC_CD_2.AddItem(" ", 0)
'            Call cob_MLT_PROC_CD_2.AddItem("VD", 1)
'            cob_MLT_PROC_CD_2.ListIndex = 0
'            CD = "BH1"
'
'        Case Else
'            cob_MLT_PROC_CD_2.Clear
'            Call cob_MLT_PROC_CD_2.AddItem(" ", 0)
'            CD = "   "
'    End Select
'    txt_MLT_PROC_CD.Text = Trim(CD)
'    End With
'End Sub

Private Sub cob_MLT_PROC_CD_1_Click()

    Dim CD As String
    
    With cob_MLT_PROC_CD_1
    
        If cob_MLT_PROC_CD_1.Text = "VD" Then
            CD = "BE1"
        ElseIf cob_MLT_PROC_CD_1.Text = "1# LF" Then
            CD = "BD1"
        ElseIf cob_MLT_PROC_CD_1.Text = "2# LF" Then
            CD = "BD2"
        ElseIf cob_MLT_PROC_CD_1.Text = "3# LF" Then
            CD = "BD3"
        ElseIf cob_MLT_PROC_CD_1.Text = "RH" Then
            CD = "BH2"
        End If
        
        txt_MLT_PROC_CD.Text = Trim(CD)
        
    End With
    
End Sub

Private Sub cob_MLT_PROC_CD_2_Change()

    Dim CD As String
    
    If cob_MLT_PROC_CD_2.Text = "1# LF" Then
        CD = "BD1"
    ElseIf cob_MLT_PROC_CD_2.Text = "2# LF" Then
        CD = "BD2"
    ElseIf cob_MLT_PROC_CD_2.Text = "3# LF" Then
        CD = "BD3"
    ElseIf cob_MLT_PROC_CD_2.Text = "VD" Then
        CD = "BE1"
    ElseIf cob_MLT_PROC_CD_2.Text = "RH" Then
        CD = "BH2"
    Else
        CD = "   "
    End If
    
    txt_MLT_PROC_CD2.Text = Trim(CD)
    
End Sub

Private Sub cob_MLT_PROC_CD_2_Click()

    Dim CD As String
    
    With cob_MLT_PROC_CD_2
    
        If .Text = "VD" Then
            CD = "BE1"
        ElseIf .Text = "1# LF" Then
            CD = "BD1"
        ElseIf .Text = "2# LF" Then
            CD = "BD2"
        ElseIf .Text = "3# LF" Then
            CD = "BD3"
        ElseIf .Text = "RH" Then
            CD = "BH2"
        End If
            
        txt_MLT_PROC_CD2.Text = Trim(CD)
        
    End With

End Sub

'Private Sub cob_MLT_PROC_CD_2_Click()
'    Dim CD As String
'
'    With cob_MLT_PROC_CD_2
'    Select Case .ListIndex
'        Case 0
'            cob_MLT_PROC_CD_3.Clear
'            Call cob_MLT_PROC_CD_3.AddItem(" ", 0)
'            cob_MLT_PROC_CD_3.ListIndex = 0
'            CD = "   "
'        Case 1
'            cob_MLT_PROC_CD_3.Clear
'            Call cob_MLT_PROC_CD_3.AddItem(" ", 0)
'            If cob_MLT_PROC_CD_2.ListCount = 5 Then
'                cob_MLT_PROC_CD_3.Clear
'                Call cob_MLT_PROC_CD_3.AddItem(" ", 0)
'                Call cob_MLT_PROC_CD_3.AddItem("VD", 1)
'                Call cob_MLT_PROC_CD_3.AddItem("RH", 2)
'                cob_MLT_PROC_CD_3.ListIndex = 0
'            ElseIf cob_MLT_PROC_CD_2.ListCount = 2 Then
'                Call cob_MLT_PROC_CD_3.AddItem("1# LF", 1)
'                Call cob_MLT_PROC_CD_3.AddItem("2# LF", 2)
'            Else
'                Call cob_MLT_PROC_CD_3.AddItem("RH", 1)
'            End If
'
'            If cob_MLT_PROC_CD_2.Text = "VD" Then
'                CD = "BE1"
'            ElseIf cob_MLT_PROC_CD_2.Text = "1# LF" Then
'                CD = "BD1"
'            Else
'                CD = "BH1"
'            End If
'            cob_MLT_PROC_CD_3.ListIndex = 0
'        Case 2
'            If cob_MLT_PROC_CD_2.ListCount = 4 Then
'                cob_MLT_PROC_CD_2.Clear
'                Call cob_MLT_PROC_CD_3.AddItem(" ", 0)
'                Call cob_MLT_PROC_CD_3.AddItem("VD", 1)
'                Call cob_MLT_PROC_CD_3.AddItem("RH", 2)
'                cob_MLT_PROC_CD_3.ListIndex = 0
'            ElseIf cob_MLT_PROC_CD_2.ListCount = 2 Then
'                cob_MLT_PROC_CD_3.Clear
'                Call cob_MLT_PROC_CD_3.AddItem(" ", 0)
'                Call cob_MLT_PROC_CD_3.AddItem("VD", 1)
'                cob_MLT_PROC_CD_3.ListIndex = 0
'            End If
'            If cob_MLT_PROC_CD_2.Text = "2# LF" Then
'                CD = "BD2"
'            Else
'                CD = "BH1"
'            End If
'        Case Else
'            cob_MLT_PROC_CD_3.Clear
'            Call cob_MLT_PROC_CD_3.AddItem(" ", 0)
'            cob_MLT_PROC_CD_3.ListIndex = 0
'            CD = "  "
'    End Select
'    txt_MLT_PROC_CD2.Text = Trim(CD)
'    End With
'End Sub

Private Sub cob_MLT_PROC_CD_3_Change()

    Dim CD As String
    
    If cob_MLT_PROC_CD_3.Text = "1# LF" Then
        CD = "BD1"
    ElseIf cob_MLT_PROC_CD_3.Text = "2# LF" Then
        CD = "BD2"
    ElseIf cob_MLT_PROC_CD_3.Text = "3# LF" Then
        CD = "BD3"
    ElseIf cob_MLT_PROC_CD_3.Text = "VD" Then
        CD = "BE1"
    ElseIf cob_MLT_PROC_CD_3.Text = "RH" Then
        CD = "BH1"
    Else
        CD = "   "
    End If
    
    txt_MLT_PROC_CD3.Text = Trim(CD)

End Sub

Private Sub cob_MLT_PROC_CD_3_Click()

    Dim CD As String
        
    With cob_MLT_PROC_CD_3
    
        If .Text = "VD" Then
            CD = "BE1"
        ElseIf .Text = "RH" Then
            CD = "BH1"
        ElseIf .Text = "1# LF" Then
            CD = "BD1"
        ElseIf .Text = "2# LF" Then
            CD = "BD2"
        ElseIf .Text = "3# LF" Then
            CD = "BD3"
        Else
            CD = "   "
        End If
        
        txt_MLT_PROC_CD3.Text = Trim(CD)
        
    End With

End Sub

Private Sub Form_Activate()
    
    'Call Gp_Combo_Add(Active_Spread)
    Call Prc_Line_Init

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

Private Sub Proc_CD_Change(ByVal sPname As Variant, ByVal iRow As Long)

On Error GoTo Process_Exec_ERROR

    Dim OutParam(2, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iCount As Integer
    
    Dim adoCmd As ADODB.Command
    
    'If ss1.MaxRows = 0 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "E_CODE"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 2
    
    OutParam(2, 1) = "arg_e_msg"
    OutParam(2, 2) = adVarChar
    OutParam(2, 3) = adParamOutput
    OutParam(2, 4) = 256
    
    sQuery = "{call AEC1070C.P_SMODIFY (  'U', "
    
    With sPname
    
        .Row = iRow
        
        .Col = 1: sQuery = sQuery + "'" + Trim(.Text) + "',"
        
        .Col = 2: sQuery = sQuery + "'" + Trim(.Text) + "',"
        
        .Col = 3: sQuery = sQuery + "'" + Trim(.Text) + "',"
        
        .Col = 8: sQuery = sQuery + "'" + Trim(.Text) + "',"
        
        .Col = 9: sQuery = sQuery + "'" + Trim(.Text) + "',?,?)}"
        
    End With
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
    
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(2, 1), OutParam(2, 2), OutParam(2, 3), OutParam(2, 4))
    
    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    If Trim(adoCmd("E_CODE")) <> "0" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
    
Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_Error : " & Error)
    
End Sub
Private Sub Prc_Line_Init()

    'txt_MLT_PROC_CD_ORG
    cob_MLT_PROC_CD_1.Clear
    Call cob_MLT_PROC_CD_1.AddItem("  ", 0)
    cob_MLT_PROC_CD_2.Clear
    Call cob_MLT_PROC_CD_2.AddItem("  ", 0)
    cob_MLT_PROC_CD_3.Clear
    Call cob_MLT_PROC_CD_3.AddItem("  ", 0)
    
    If Mid(txt_MLT_PROC_CD_ORG.Text, 4, 3) = "BD1" Then
        Call cob_MLT_PROC_CD_1.AddItem("2# LF", 1)
        Call cob_MLT_PROC_CD_1.AddItem("3# LF", 2)
    ElseIf Mid(txt_MLT_PROC_CD_ORG.Text, 4, 3) = "BD2" Then
        Call cob_MLT_PROC_CD_1.AddItem("1# LF", 1)
        Call cob_MLT_PROC_CD_1.AddItem("3# LF", 2)
    ElseIf Mid(txt_MLT_PROC_CD_ORG.Text, 4, 3) = "BD3" Then
        Call cob_MLT_PROC_CD_1.AddItem("1# LF", 1)
        Call cob_MLT_PROC_CD_1.AddItem("2# LF", 2)
    ElseIf Mid(txt_MLT_PROC_CD_ORG.Text, 4, 3) = "BE1" Then
        Call cob_MLT_PROC_CD_1.AddItem("RH", 1)
    ElseIf Mid(txt_MLT_PROC_CD_ORG.Text, 4, 3) = "BH2" Then
        Call cob_MLT_PROC_CD_1.AddItem("VD", 1)
    End If

    If cob_MLT_PROC_CD_1.ListCount > 1 Then
        If Mid(txt_MLT_PROC_CD_ORG.Text, 7, 3) = "BD1" Then
            Call cob_MLT_PROC_CD_2.AddItem("2# LF", 1)
            Call cob_MLT_PROC_CD_2.AddItem("3# LF", 2)
        ElseIf Mid(txt_MLT_PROC_CD_ORG.Text, 7, 3) = "BD2" Then
            Call cob_MLT_PROC_CD_2.AddItem("1# LF", 1)
            Call cob_MLT_PROC_CD_2.AddItem("3# LF", 2)
        ElseIf Mid(txt_MLT_PROC_CD_ORG.Text, 7, 3) = "BD3" Then
            Call cob_MLT_PROC_CD_2.AddItem("1# LF", 1)
            Call cob_MLT_PROC_CD_2.AddItem("2# LF", 2)
        ElseIf Mid(txt_MLT_PROC_CD_ORG.Text, 7, 3) = "BE1" Then
            Call cob_MLT_PROC_CD_2.AddItem("RH", 1)
        ElseIf Mid(txt_MLT_PROC_CD_ORG.Text, 7, 3) = "BH2" Then
            Call cob_MLT_PROC_CD_2.AddItem("VD", 1)
        End If
    Else
        If Mid(txt_MLT_PROC_CD_ORG.Text, 7, 3) = "BD1" Then
            Call cob_MLT_PROC_CD_1.AddItem("2# LF", 1)
            Call cob_MLT_PROC_CD_1.AddItem("3# LF", 2)
        ElseIf Mid(txt_MLT_PROC_CD_ORG.Text, 7, 3) = "BD2" Then
            Call cob_MLT_PROC_CD_1.AddItem("1# LF", 1)
            Call cob_MLT_PROC_CD_1.AddItem("3# LF", 2)
        ElseIf Mid(txt_MLT_PROC_CD_ORG.Text, 7, 3) = "BD3" Then
            Call cob_MLT_PROC_CD_1.AddItem("1# LF", 1)
            Call cob_MLT_PROC_CD_1.AddItem("2# LF", 2)
        ElseIf Mid(txt_MLT_PROC_CD_ORG.Text, 7, 3) = "BE1" Then
            Call cob_MLT_PROC_CD_1.AddItem("RH", 1)
        ElseIf Mid(txt_MLT_PROC_CD_ORG.Text, 7, 3) = "BH2" Then
            Call cob_MLT_PROC_CD_1.AddItem("VD", 1)
        End If
    End If
    
    If Mid(txt_MLT_PROC_CD_ORG.Text, 9, 3) = "BD1" Then
        Call cob_MLT_PROC_CD_2.AddItem("2# LF", 1)
        Call cob_MLT_PROC_CD_2.AddItem("3# LF", 2)
    ElseIf Mid(txt_MLT_PROC_CD_ORG.Text, 9, 3) = "BD2" Then
        Call cob_MLT_PROC_CD_2.AddItem("1# LF", 1)
        Call cob_MLT_PROC_CD_2.AddItem("3# LF", 2)
    ElseIf Mid(txt_MLT_PROC_CD_ORG.Text, 9, 3) = "BD3" Then
        Call cob_MLT_PROC_CD_2.AddItem("1# LF", 1)
        Call cob_MLT_PROC_CD_2.AddItem("2# LF", 2)
    ElseIf Mid(txt_MLT_PROC_CD_ORG.Text, 9, 3) = "BE1" Then
        Call cob_MLT_PROC_CD_2.AddItem("RH", 1)
    ElseIf Mid(txt_MLT_PROC_CD_ORG.Text, 9, 3) = "BH2" Then
        Call cob_MLT_PROC_CD_2.AddItem("VD", 1)
    End If

End Sub
